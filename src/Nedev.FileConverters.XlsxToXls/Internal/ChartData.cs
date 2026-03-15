namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Represents chart data for BIFF8 format conversion.
/// Contains all information needed to render an Excel chart in XLS format.
/// </summary>
public sealed class ChartData
{
    /// <summary>Gets or sets the chart name.</summary>
    public string Name { get; set; } = "Chart";

    /// <summary>Gets or sets the chart type.</summary>
    public ChartType Type { get; set; } = ChartType.Column;

    /// <summary>Gets or sets the chart position and size.</summary>
    public ChartPosition Position { get; set; } = new();

    /// <summary>Gets or sets the data series collection.</summary>
    public List<ChartSeries> Series { get; set; } = [];

    /// <summary>Gets or sets the chart title.</summary>
    public ChartTitle? Title { get; set; }

    /// <summary>Gets or sets the category (X) axis.</summary>
    public ChartAxis? CategoryAxis { get; set; }

    /// <summary>Gets or sets the value (Y) axis.</summary>
    public ChartAxis? ValueAxis { get; set; }

    /// <summary>Gets or sets the chart legend.</summary>
    public ChartLegend? Legend { get; set; }

    /// <summary>Gets or sets the plot area configuration.</summary>
    public ChartPlotArea PlotArea { get; set; } = new();
}

/// <summary>Supported chart types for BIFF8 format.</summary>
public enum ChartType : ushort
{
    /// <summary>Area chart.</summary>
    Area = 0x0001,
    /// <summary>Bar chart.</summary>
    Bar = 0x0002,
    /// <summary>Line chart.</summary>
    Line = 0x0003,
    /// <summary>Pie chart.</summary>
    Pie = 0x0004,
    /// <summary>Scatter chart.</summary>
    Scatter = 0x0005,
    /// <summary>Radar chart.</summary>
    Radar = 0x0006,
    /// <summary>Column chart (default).</summary>
    Column = 0x0008,
    /// <summary>Doughnut chart.</summary>
    Doughnut = 0x0009
}

/// <summary>Defines the position and dimensions of a chart.</summary>
public sealed class ChartPosition
{
    /// <summary>Gets or sets the X coordinate in points.</summary>
    public int X { get; set; }

    /// <summary>Gets or sets the Y coordinate in points.</summary>
    public int Y { get; set; }

    /// <summary>Gets or sets the width in points. Default is 400.</summary>
    public int Width { get; set; } = 400;

    /// <summary>Gets or sets the height in points. Default is 300.</summary>
    public int Height { get; set; } = 300;

    /// <summary>Gets or sets whether the position is absolute.</summary>
    public bool IsAbsolute { get; set; }
}

/// <summary>Represents a data series in a chart.</summary>
public sealed class ChartSeries
{
    /// <summary>Gets or sets the series name.</summary>
    public string? Name { get; set; }

    /// <summary>Gets or sets the category data range.</summary>
    public ChartRange? Categories { get; set; }

    /// <summary>Gets or sets the value data range.</summary>
    public ChartRange? Values { get; set; }

    /// <summary>Gets or sets the series formula.</summary>
    public string? Formula { get; set; }

    /// <summary>Gets or sets the series index.</summary>
    public int SeriesIndex { get; set; }

    /// <summary>Gets or sets the category index.</summary>
    public int CategoryIndex { get; set; }

    /// <summary>Gets or sets the value index.</summary>
    public int ValueIndex { get; set; }

    /// <summary>Gets or sets the bubble index. Default is -1.</summary>
    public int BubbleIndex { get; set; } = -1;

    /// <summary>Gets or sets the data labels configuration.</summary>
    public DataLabels? DataLabels { get; set; }

    /// <summary>Gets or sets the fill color.</summary>
    public ChartColor? FillColor { get; set; }

    /// <summary>Gets or sets the border color.</summary>
    public ChartColor? BorderColor { get; set; }

    /// <summary>Gets or sets the line style.</summary>
    public LineStyle? LineStyle { get; set; }

    /// <summary>Gets or sets the marker style. Default is None.</summary>
    public MarkerStyle MarkerStyle { get; set; } = MarkerStyle.None;

    /// <summary>Gets or sets individual data point styles.</summary>
    public List<ChartDataPoint>? DataPoints { get; set; }

    /// <summary>Gets or sets trend lines for this series.</summary>
    public List<TrendLine>? TrendLines { get; set; }

    /// <summary>Gets or sets error bars for this series.</summary>
    public ErrorBars? ErrorBars { get; set; }

    /// <summary>Gets or sets the secondary chart type for combo charts.</summary>
    public ChartType? SecondaryChartType { get; set; }

    /// <summary>Gets or sets whether to use the secondary axis.</summary>
    public bool UseSecondaryAxis { get; set; }
}

/// <summary>Represents an individual data point with custom styling.</summary>
public sealed class ChartDataPoint
{
    /// <summary>Gets or sets the data point index.</summary>
    public int Index { get; set; }

    /// <summary>Gets or sets the fill color.</summary>
    public ChartColor? FillColor { get; set; }

    /// <summary>Gets or sets the border color.</summary>
    public ChartColor? BorderColor { get; set; }

    /// <summary>Gets or sets the data labels for this point.</summary>
    public DataLabels? DataLabels { get; set; }

    /// <summary>Gets or sets whether to explode this point (for pie charts).</summary>
    public bool? Explosion { get; set; }
}

/// <summary>Represents a trend line for data analysis.</summary>
public sealed class TrendLine
{
    /// <summary>Gets or sets the trend line type. Default is Linear.</summary>
    public TrendLineType Type { get; set; } = TrendLineType.Linear;

    /// <summary>Gets or sets the trend line name.</summary>
    public string? Name { get; set; }

    /// <summary>Gets or sets whether to display the equation.</summary>
    public bool DisplayEquation { get; set; }

    /// <summary>Gets or sets whether to display the R-squared value.</summary>
    public bool DisplayRSquared { get; set; }

    /// <summary>Gets or sets the line color.</summary>
    public ChartColor? LineColor { get; set; }

    /// <summary>Gets or sets the line style. Default is Solid.</summary>
    public LineStyle LineStyle { get; set; } = LineStyle.Solid;

    /// <summary>Gets or sets the polynomial order. Default is 2.</summary>
    public int Order { get; set; } = 2;

    /// <summary>Gets or sets the moving average period. Default is 2.</summary>
    public int Period { get; set; } = 2;

    /// <summary>Gets or sets the forward forecast period.</summary>
    public double? Forward { get; set; }

    /// <summary>Gets or sets the backward forecast period.</summary>
    public double? Backward { get; set; }
}

/// <summary>Types of trend lines for data analysis.</summary>
public enum TrendLineType : byte
{
    /// <summary>Linear trend line.</summary>
    Linear = 0,
    /// <summary>Exponential trend line.</summary>
    Exponential = 1,
    /// <summary>Logarithmic trend line.</summary>
    Logarithmic = 2,
    /// <summary>Polynomial trend line.</summary>
    Polynomial = 3,
    /// <summary>Power trend line.</summary>
    Power = 4,
    /// <summary>Moving average trend line.</summary>
    MovingAverage = 5
}

/// <summary>Represents error bars for data series.</summary>
public sealed class ErrorBars
{
    /// <summary>Gets or sets the error bar type. Default is Both.</summary>
    public ErrorBarType Type { get; set; } = ErrorBarType.Both;

    /// <summary>Gets or sets the value type for error bars. Default is FixedValue.</summary>
    public ErrorBarValueType ValueType { get; set; } = ErrorBarValueType.FixedValue;

    /// <summary>Gets or sets the error value.</summary>
    public double Value { get; set; }

    /// <summary>Gets or sets whether to show end caps. Default is true.</summary>
    public bool ShowCap { get; set; } = true;

    /// <summary>Gets or sets the line color.</summary>
    public ChartColor? LineColor { get; set; }

    /// <summary>Gets or sets the line style. Default is Solid.</summary>
    public LineStyle LineStyle { get; set; } = LineStyle.Solid;
}

/// <summary>Error bar direction types.</summary>
public enum ErrorBarType : byte
{
    /// <summary>Error bars in both directions.</summary>
    Both = 0,
    /// <summary>Error bars in positive direction only.</summary>
    Plus = 1,
    /// <summary>Error bars in negative direction only.</summary>
    Minus = 2
}

/// <summary>Error bar value calculation types.</summary>
public enum ErrorBarValueType : byte
{
    /// <summary>Fixed value.</summary>
    FixedValue = 0,
    /// <summary>Percentage of the value.</summary>
    Percentage = 1,
    /// <summary>Standard deviation.</summary>
    StandardDeviation = 2,
    /// <summary>Standard error.</summary>
    StandardError = 3,
    /// <summary>Custom values.</summary>
    Custom = 4
}

/// <summary>Configuration for data labels.</summary>
public sealed class DataLabels
{
    /// <summary>Gets or sets whether to show data labels. Default is true.</summary>
    public bool Show { get; set; } = true;

    /// <summary>Gets or sets whether to show values. Default is true.</summary>
    public bool ShowValue { get; set; } = true;

    /// <summary>Gets or sets whether to show category names.</summary>
    public bool ShowCategory { get; set; }

    /// <summary>Gets or sets whether to show percentages.</summary>
    public bool ShowPercentage { get; set; }

    /// <summary>Gets or sets whether to show series names.</summary>
    public bool ShowSeriesName { get; set; }

    /// <summary>Gets or sets the label position. Default is OutsideEnd.</summary>
    public DataLabelPosition Position { get; set; } = DataLabelPosition.OutsideEnd;
}

/// <summary>Positions for data labels.</summary>
public enum DataLabelPosition : byte
{
    /// <summary>Center of the data point.</summary>
    Center = 0,
    /// <summary>Inside end of the data point.</summary>
    InsideEnd = 1,
    /// <summary>Outside end of the data point.</summary>
    OutsideEnd = 2,
    /// <summary>Best fit position.</summary>
    BestFit = 3,
    /// <summary>Left side.</summary>
    Left = 4,
    /// <summary>Right side.</summary>
    Right = 5,
    /// <summary>Above the data point.</summary>
    Above = 6,
    /// <summary>Below the data point.</summary>
    Below = 7
}

/// <summary>Marker styles for line and scatter charts.</summary>
public enum MarkerStyle : byte
{
    /// <summary>No marker.</summary>
    None = 0,
    /// <summary>Square marker.</summary>
    Square = 1,
    /// <summary>Diamond marker.</summary>
    Diamond = 2,
    /// <summary>Triangle marker.</summary>
    Triangle = 3,
    /// <summary>X marker.</summary>
    X = 4,
    /// <summary>Star marker.</summary>
    Star = 5,
    /// <summary>Dot marker.</summary>
    Dot = 6,
    /// <summary>Circle marker.</summary>
    Circle = 7,
    /// <summary>Plus marker.</summary>
    Plus = 8
}

/// <summary>Line styles for chart elements.</summary>
public enum LineStyle : byte
{
    /// <summary>Solid line.</summary>
    Solid = 0,
    /// <summary>Dashed line.</summary>
    Dash = 1,
    /// <summary>Dotted line.</summary>
    Dot = 2,
    /// <summary>Dash-dot line.</summary>
    DashDot = 3,
    /// <summary>Dash-dot-dot line.</summary>
    DashDotDot = 4,
    /// <summary>No line.</summary>
    None = 5
}

/// <summary>Represents an RGB color for chart elements.</summary>
/// <param name="R">Red component (0-255).</param>
/// <param name="G">Green component (0-255).</param>
/// <param name="B">Blue component (0-255).</param>
public readonly record struct ChartColor(byte R, byte G, byte B)
{
    /// <summary>Creates a color from RGB values.</summary>
    public static ChartColor FromRgb(byte r, byte g, byte b) => new(r, g, b);

    /// <summary>Red color.</summary>
    public static readonly ChartColor Red = new(255, 0, 0);
    /// <summary>Green color.</summary>
    public static readonly ChartColor Green = new(0, 255, 0);
    /// <summary>Blue color.</summary>
    public static readonly ChartColor Blue = new(0, 0, 255);
    /// <summary>Yellow color.</summary>
    public static readonly ChartColor Yellow = new(255, 255, 0);
    /// <summary>Cyan color.</summary>
    public static readonly ChartColor Cyan = new(0, 255, 255);
    /// <summary>Magenta color.</summary>
    public static readonly ChartColor Magenta = new(255, 0, 255);
    /// <summary>Black color.</summary>
    public static readonly ChartColor Black = new(0, 0, 0);
    /// <summary>White color.</summary>
    public static readonly ChartColor White = new(255, 255, 255);
    /// <summary>Gray color.</summary>
    public static readonly ChartColor Gray = new(128, 128, 128);
    /// <summary>Orange color.</summary>
    public static readonly ChartColor Orange = new(255, 165, 0);
    /// <summary>Purple color.</summary>
    public static readonly ChartColor Purple = new(128, 0, 128);
    /// <summary>Dark red color.</summary>
    public static readonly ChartColor DarkRed = new(139, 0, 0);
    /// <summary>Dark green color.</summary>
    public static readonly ChartColor DarkGreen = new(0, 100, 0);
    /// <summary>Dark blue color.</summary>
    public static readonly ChartColor DarkBlue = new(0, 0, 139);
}

/// <summary>Represents a cell range reference.</summary>
public sealed class ChartRange
{
    /// <summary>Gets or sets the sheet name.</summary>
    public string SheetName { get; init; } = "";

    /// <summary>Gets or sets the first row index (0-based).</summary>
    public int FirstRow { get; init; }

    /// <summary>Gets or sets the first column index (0-based).</summary>
    public int FirstCol { get; init; }

    /// <summary>Gets or sets the last row index (0-based).</summary>
    public int LastRow { get; init; }

    /// <summary>Gets or sets the last column index (0-based).</summary>
    public int LastCol { get; init; }

    /// <summary>Gets whether the range is a single cell.</summary>
    public bool IsSingleCell => FirstRow == LastRow && FirstCol == LastCol;

    /// <summary>Creates a new range with a different sheet name.</summary>
    public ChartRange WithSheetName(string name) => new() { SheetName = name, FirstRow = FirstRow, FirstCol = FirstCol, LastRow = LastRow, LastCol = LastCol };

    /// <summary>Creates a new range with different row boundaries.</summary>
    public ChartRange WithRows(int firstRow, int lastRow) => new() { SheetName = SheetName, FirstRow = firstRow, FirstCol = FirstCol, LastRow = lastRow, LastCol = LastCol };

    /// <summary>Creates a new range with different column boundaries.</summary>
    public ChartRange WithCols(int firstCol, int lastCol) => new() { SheetName = SheetName, FirstRow = FirstRow, FirstCol = firstCol, LastRow = LastRow, LastCol = lastCol };

    /// <summary>Creates a new range for a single cell.</summary>
    public ChartRange WithCell(int row, int col) => new() { SheetName = SheetName, FirstRow = row, FirstCol = col, LastRow = row, LastCol = col };
}

/// <summary>Represents a chart title.</summary>
public sealed class ChartTitle
{
    /// <summary>Gets or sets the title text.</summary>
    public string Text { get; set; } = "";

    /// <summary>Gets or sets the title position.</summary>
    public ChartPosition Position { get; set; } = new();
}

/// <summary>Represents a chart axis.</summary>
public sealed class ChartAxis
{
    /// <summary>Gets or sets the axis type. Default is Category.</summary>
    public AxisType Type { get; set; } = AxisType.Category;

    /// <summary>Gets or sets the axis position. Default is Bottom.</summary>
    public AxisPosition Position { get; set; } = AxisPosition.Bottom;

    /// <summary>Gets or sets the axis title.</summary>
    public string? Title { get; set; }

    /// <summary>Gets or sets the minimum value.</summary>
    public double? MinValue { get; set; }

    /// <summary>Gets or sets the maximum value.</summary>
    public double? MaxValue { get; set; }

    /// <summary>Gets or sets whether to show major gridlines. Default is true.</summary>
    public bool HasMajorGridlines { get; set; } = true;

    /// <summary>Gets or sets whether to show minor gridlines.</summary>
    public bool HasMinorGridlines { get; set; }
}

/// <summary>Types of chart axes.</summary>
public enum AxisType : byte
{
    /// <summary>Category axis (X-axis).</summary>
    Category = 0,
    /// <summary>Value axis (Y-axis).</summary>
    Value = 1,
    /// <summary>Series axis (for 3D charts).</summary>
    Series = 2
}

/// <summary>Positions for chart axes.</summary>
public enum AxisPosition : byte
{
    /// <summary>Bottom position.</summary>
    Bottom = 0,
    /// <summary>Left position.</summary>
    Left = 1,
    /// <summary>Top position.</summary>
    Top = 2,
    /// <summary>Right position.</summary>
    Right = 3
}

/// <summary>Represents a chart legend.</summary>
public sealed class ChartLegend
{
    /// <summary>Gets or sets the legend position. Default is Right.</summary>
    public LegendPosition Position { get; set; } = LegendPosition.Right;

    /// <summary>Gets or sets whether to show the legend. Default is true.</summary>
    public bool Show { get; set; } = true;
}

/// <summary>Positions for the chart legend.</summary>
public enum LegendPosition : byte
{
    /// <summary>Right side of the chart.</summary>
    Right = 0,
    /// <summary>Left side of the chart.</summary>
    Left = 1,
    /// <summary>Bottom of the chart.</summary>
    Bottom = 2,
    /// <summary>Top of the chart.</summary>
    Top = 3,
    /// <summary>Corner of the chart.</summary>
    Corner = 4
}

/// <summary>Represents the chart plot area.</summary>
public sealed class ChartPlotArea
{
    /// <summary>Gets or sets the X position. Default is 20.</summary>
    public int X { get; set; } = 20;

    /// <summary>Gets or sets the Y position. Default is 20.</summary>
    public int Y { get; set; } = 20;

    /// <summary>Gets or sets the width. Default is 360.</summary>
    public int Width { get; set; } = 360;

    /// <summary>Gets or sets the height. Default is 240.</summary>
    public int Height { get; set; } = 240;

    /// <summary>Gets or sets whether to vary colors by data point.</summary>
    public bool VaryColors { get; set; }
}
