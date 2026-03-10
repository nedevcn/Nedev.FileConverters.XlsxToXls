using System.IO.Compression;
using System.Xml;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// XLSX图表读取器 - 从Open XML格式读取图表数据
/// </summary>
internal static class ChartReader
{
    private static readonly XmlReaderSettings Settings = new()
    {
        Async = false,
        CloseInput = false,
        IgnoreWhitespace = true,
        DtdProcessing = DtdProcessing.Prohibit
    };

    public static List<ChartData> ReadCharts(ZipArchive archive, string worksheetPath)
    {
        var charts = new List<ChartData>();
        var drawingPath = FindDrawingPath(archive, worksheetPath);
        if (string.IsNullOrEmpty(drawingPath)) return charts;

        var chartRefs = ReadDrawingForCharts(archive, drawingPath);
        foreach (var chartPath in chartRefs)
        {
            var chart = ReadChartFile(archive, chartPath);
            if (chart != null) charts.Add(chart);
        }
        return charts;
    }

    private static string? FindDrawingPath(ZipArchive archive, string worksheetPath)
    {
        var relsPath = worksheetPath.Replace("worksheets/", "worksheets/_rels/") + ".rels";
        var entry = archive.GetEntry(relsPath) ?? archive.GetEntry("xl/" + relsPath);
        if (entry == null) return null;

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Settings);
        var ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship" && reader.NamespaceURI == ns)
            {
                var type = reader.GetAttribute("Type");
                var target = reader.GetAttribute("Target");
                if (type?.EndsWith("/drawing", StringComparison.OrdinalIgnoreCase) == true && target != null)
                {
                    return target.StartsWith("/") ? target.TrimStart('/') : "xl/drawings/" + target.Split('/').Last();
                }
            }
        }
        return null;
    }

    private static List<string> ReadDrawingForCharts(ZipArchive archive, string drawingPath)
    {
        var chartPaths = new List<string>();
        var entry = archive.GetEntry(drawingPath) ?? archive.GetEntry("xl/drawings/" + drawingPath.Split('/').Last());
        if (entry == null) return chartPaths;

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Settings);
        var chartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "chart" && reader.NamespaceURI == chartNs)
                {
                var id = reader.GetAttribute("id");
                if (!string.IsNullOrEmpty(id))
                {
                    var chartPath = ResolveChartPath(archive, drawingPath, id);
                    if (!string.IsNullOrEmpty(chartPath)) chartPaths.Add(chartPath);
                }
            }
        }
        return chartPaths;
    }

    private static string? ResolveChartPath(ZipArchive archive, string drawingPath, string rId)
    {
        var relsPath = drawingPath.Replace("drawings/", "drawings/_rels/") + ".rels";
        var entry = archive.GetEntry(relsPath) ?? archive.GetEntry("xl/" + relsPath);
        if (entry == null) return null;

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Settings);
        var ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship" && reader.NamespaceURI == ns)
            {
                var id = reader.GetAttribute("Id");
                var type = reader.GetAttribute("Type");
                var target = reader.GetAttribute("Target");
                if (id == rId && type?.EndsWith("/chart", StringComparison.OrdinalIgnoreCase) == true && target != null)
                {
                    return target.StartsWith("/") ? target.TrimStart('/') : "xl/charts/" + target.Split('/').Last();
                }
            }
        }
        return null;
    }

    private static ChartData? ReadChartFile(ZipArchive archive, string chartPath)
    {
        var entry = archive.GetEntry(chartPath) ?? archive.GetEntry("xl/charts/" + chartPath.Split('/').Last());
        if (entry == null) return null;

        var chart = new ChartData();
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Settings);
        var ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element) continue;

            switch (reader.LocalName)
            {
                case "barChart":
                    chart.Type = ChartType.Bar;
                    ReadBarChart(reader, ns, chart);
                    break;
                case "lineChart":
                    chart.Type = ChartType.Line;
                    ReadLineChart(reader, ns, chart);
                    break;
                case "pieChart":
                case "pie3DChart":
                    chart.Type = ChartType.Pie;
                    ReadPieChart(reader, ns, chart);
                    break;
                case "areaChart":
                    chart.Type = ChartType.Area;
                    ReadAreaChart(reader, ns, chart);
                    break;
                case "scatterChart":
                    chart.Type = ChartType.Scatter;
                    ReadScatterChart(reader, ns, chart);
                    break;
                case "radarChart":
                    chart.Type = ChartType.Radar;
                    ReadRadarChart(reader, ns, chart);
                    break;
                case "doughnutChart":
                    chart.Type = ChartType.Doughnut;
                    ReadDoughnutChart(reader, ns, chart);
                    break;
                case "title":
                    chart.Title = ReadTitle(reader, ns);
                    break;
                case "legend":
                    chart.Legend = ReadLegend(reader, ns);
                    break;
                case "catAx":
                    chart.CategoryAxis = ReadAxis(reader, ns, AxisType.Category);
                    break;
                case "valAx":
                    chart.ValueAxis = ReadAxis(reader, ns, AxisType.Value);
                    break;
            }
        }

        return chart.Series.Count > 0 ? chart : null;
    }

    private static void ReadBarChart(XmlReader reader, string ns, ChartData chart)
    {
        if (reader.IsEmptyElement) return;
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                depth++;
                if (reader.LocalName == "ser" && reader.NamespaceURI == ns)
                {
                    var series = ReadSeries(reader, ns);
                    if (series != null) chart.Series.Add(series);
                    depth--;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static void ReadLineChart(XmlReader reader, string ns, ChartData chart)
    {
        ReadBarChart(reader, ns, chart);
    }

    private static void ReadPieChart(XmlReader reader, string ns, ChartData chart)
    {
        ReadBarChart(reader, ns, chart);
    }

    private static void ReadAreaChart(XmlReader reader, string ns, ChartData chart)
    {
        ReadBarChart(reader, ns, chart);
    }

    private static void ReadScatterChart(XmlReader reader, string ns, ChartData chart)
    {
        ReadBarChart(reader, ns, chart);
    }

    private static void ReadRadarChart(XmlReader reader, string ns, ChartData chart)
    {
        ReadBarChart(reader, ns, chart);
    }

    private static void ReadDoughnutChart(XmlReader reader, string ns, ChartData chart)
    {
        ReadBarChart(reader, ns, chart);
    }

    private static ChartSeries? ReadSeries(XmlReader reader, string ns)
    {
        var series = new ChartSeries();
        if (reader.IsEmptyElement) return series;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                depth++;
                switch (reader.LocalName)
                {
                    case "tx":
                        series.Name = ReadSeriesText(reader, ns);
                        depth--;
                        break;
                    case "cat":
                        case "xVal":
                        series.Categories = ReadChartReference(reader, ns);
                        depth--;
                        break;
                    case "val":
                    case "yVal":
                        series.Values = ReadChartReference(reader, ns);
                        depth--;
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return series;
    }

    private static string? ReadSeriesText(XmlReader reader, string ns)
    {
        if (reader.IsEmptyElement) return null;
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "v")
                {
                    if (reader.Read() && reader.NodeType == XmlNodeType.Text)
                        return reader.Value;
                }
                else if (reader.LocalName == "f")
                {
                    var formula = reader.ReadElementContentAsString();
                    return formula;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return null;
    }

    private static ChartRange? ReadChartReference(XmlReader reader, string ns)
    {
        if (reader.IsEmptyElement) return new ChartRange();

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "f")
                {
                    var formula = reader.ReadElementContentAsString();
                    return ParseFormula(formula);
                }
                else if (reader.LocalName == "strRef" || reader.LocalName == "numRef")
                {
                    var innerDepth = 1;
                    while (reader.Read() && innerDepth > 0)
                    {
                        if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns && reader.LocalName == "f")
                        {
                            var formula = reader.ReadElementContentAsString();
                            return ParseFormula(formula);
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement) innerDepth--;
                    }
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return new ChartRange();
    }

    private static ChartRange ParseFormula(string formula)
    {
        var sheetName = "";
        var excl = formula.IndexOf('!');
        if (excl >= 0)
        {
            sheetName = formula[..excl].Trim('\'');
            formula = formula[(excl + 1)..];
        }

        formula = formula.Replace("$", "");
        var colon = formula.IndexOf(':');
        if (colon >= 0)
        {
            var left = formula[..colon];
            var right = formula[(colon + 1)..];
            ParseCellRef(left, out var firstRow, out var firstCol);
            ParseCellRef(right, out var lastRow, out var lastCol);
            return new ChartRange { SheetName = sheetName, FirstRow = firstRow, FirstCol = firstCol, LastRow = lastRow, LastCol = lastCol };
        }
        else
        {
            ParseCellRef(formula, out var row, out var col);
            return new ChartRange { SheetName = sheetName, FirstRow = row, FirstCol = col, LastRow = row, LastCol = col };
        }
    }

    private static void ParseCellRef(string s, out int row, out int col)
    {
        row = col = 0;
        if (string.IsNullOrEmpty(s)) return;

        var i = 0;
        while (i < s.Length && char.IsLetter(s[i])) i++;
        if (i == 0 || i >= s.Length) return;

        if (!int.TryParse(s.AsSpan(i), out var r)) return;
        row = r - 1;
        col = ParseCol(s.AsSpan(0, i));
    }

    private static int ParseCol(ReadOnlySpan<char> s)
    {
        var col = 0;
        foreach (var c in s)
            col = col * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        return col - 1;
    }

    private static ChartTitle? ReadTitle(XmlReader reader, string ns)
    {
        var title = new ChartTitle();
        if (reader.IsEmptyElement) return title;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "v")
                {
                    if (reader.Read() && reader.NodeType == XmlNodeType.Text)
                        title.Text = reader.Value;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return title;
    }

    private static ChartLegend? ReadLegend(XmlReader reader, string ns)
    {
        var legend = new ChartLegend();
        var pos = reader.GetAttribute("legendPos");
        if (!string.IsNullOrEmpty(pos))
        {
            legend.Position = pos.ToLowerInvariant() switch
            {
                "l" => LegendPosition.Left,
                "r" => LegendPosition.Right,
                "t" => LegendPosition.Top,
                "b" => LegendPosition.Bottom,
                "tr" => LegendPosition.Corner,
                _ => LegendPosition.Right
            };
        }
        return legend;
    }

    private static ChartAxis ReadAxis(XmlReader reader, string ns, AxisType type)
    {
        var axis = new ChartAxis { Type = type };
        if (reader.IsEmptyElement) return axis;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "title":
                        var title = ReadTitle(reader, ns);
                        if (title != null) axis.Title = title.Text;
                        break;
                    case "scaling":
                        ReadScaling(reader, ns, axis);
                        break;
                    case "majorGridlines":
                        axis.HasMajorGridlines = true;
                        break;
                    case "minorGridlines":
                        axis.HasMinorGridlines = true;
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return axis;
    }

    private static void ReadScaling(XmlReader reader, string ns, ChartAxis axis)
    {
        if (reader.IsEmptyElement) return;
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "min")
                {
                    var val = reader.GetAttribute("val");
                    if (double.TryParse(val, out var min)) axis.MinValue = min;
                }
                else if (reader.LocalName == "max")
                {
                    var val = reader.GetAttribute("val");
                    if (double.TryParse(val, out var max)) axis.MaxValue = max;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }
}
