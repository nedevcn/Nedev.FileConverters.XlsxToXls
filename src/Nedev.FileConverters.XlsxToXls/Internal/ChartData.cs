namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// 图表数据模型 - 支持BIFF8图表格式
/// </summary>
public sealed class ChartData
{
    public string Name { get; set; } = "Chart";
    public ChartType Type { get; set; } = ChartType.Column;
    public ChartPosition Position { get; set; } = new();
    public List<ChartSeries> Series { get; set; } = [];
    public ChartTitle? Title { get; set; }
    public ChartAxis? CategoryAxis { get; set; }
    public ChartAxis? ValueAxis { get; set; }
    public ChartLegend? Legend { get; set; }
    public ChartPlotArea PlotArea { get; set; } = new();
}

public enum ChartType : ushort
{
    Area = 0x0001,
    Bar = 0x0002,
    Line = 0x0003,
    Pie = 0x0004,
    Scatter = 0x0005,
    Radar = 0x0006,
    Column = 0x0008,
    Doughnut = 0x0009
}

public sealed class ChartPosition
{
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; } = 400;
    public int Height { get; set; } = 300;
    public bool IsAbsolute { get; set; }
}

public sealed class ChartSeries
{
    public string? Name { get; set; }
    public ChartRange? Categories { get; set; }
    public ChartRange? Values { get; set; }
    public string? Formula { get; set; }
    public int SeriesIndex { get; set; }
    public int CategoryIndex { get; set; }
    public int ValueIndex { get; set; }
    public int BubbleIndex { get; set; } = -1;

    // 数据标签设置
    public DataLabels? DataLabels { get; set; }

    // 系列样式
    public ChartColor? FillColor { get; set; }
    public ChartColor? BorderColor { get; set; }
    public LineStyle? LineStyle { get; set; }
    public MarkerStyle MarkerStyle { get; set; } = MarkerStyle.None;

    // 数据点级别设置
    public List<ChartDataPoint>? DataPoints { get; set; }

    // 趋势线
    public List<TrendLine>? TrendLines { get; set; }

    // 误差线
    public ErrorBars? ErrorBars { get; set; }

    // 组合图表类型（用于组合图表）
    public ChartType? SecondaryChartType { get; set; }
    public bool UseSecondaryAxis { get; set; }
}

/// <summary>
/// 图表数据点 - 支持单个数据点的独立样式
/// </summary>
public sealed class ChartDataPoint
{
    public int Index { get; set; }
    public ChartColor? FillColor { get; set; }
    public ChartColor? BorderColor { get; set; }
    public DataLabels? DataLabels { get; set; }
    public bool? Explosion { get; set; }
}

/// <summary>
/// 趋势线
/// </summary>
public sealed class TrendLine
{
    public TrendLineType Type { get; set; } = TrendLineType.Linear;
    public string? Name { get; set; }
    public bool DisplayEquation { get; set; }
    public bool DisplayRSquared { get; set; }
    public ChartColor? LineColor { get; set; }
    public LineStyle LineStyle { get; set; } = LineStyle.Solid;
    public int Order { get; set; } = 2;
    public int Period { get; set; } = 2;
    public double? Forward { get; set; }
    public double? Backward { get; set; }
}

public enum TrendLineType : byte
{
    Linear = 0,
    Exponential = 1,
    Logarithmic = 2,
    Polynomial = 3,
    Power = 4,
    MovingAverage = 5
}

/// <summary>
/// 误差线
/// </summary>
public sealed class ErrorBars
{
    public ErrorBarType Type { get; set; } = ErrorBarType.Both;
    public ErrorBarValueType ValueType { get; set; } = ErrorBarValueType.FixedValue;
    public double Value { get; set; }
    public bool ShowCap { get; set; } = true;
    public ChartColor? LineColor { get; set; }
    public LineStyle LineStyle { get; set; } = LineStyle.Solid;
}

public enum ErrorBarType : byte
{
    Both = 0,
    Plus = 1,
    Minus = 2
}

public enum ErrorBarValueType : byte
{
    FixedValue = 0,
    Percentage = 1,
    StandardDeviation = 2,
    StandardError = 3,
    Custom = 4
}

public sealed class DataLabels
{
    public bool Show { get; set; } = true;
    public bool ShowValue { get; set; } = true;
    public bool ShowCategory { get; set; }
    public bool ShowPercentage { get; set; }
    public bool ShowSeriesName { get; set; }
    public DataLabelPosition Position { get; set; } = DataLabelPosition.OutsideEnd;
}

public enum DataLabelPosition : byte
{
    Center = 0,
    InsideEnd = 1,
    OutsideEnd = 2,
    BestFit = 3,
    Left = 4,
    Right = 5,
    Above = 6,
    Below = 7
}

public enum MarkerStyle : byte
{
    None = 0,
    Square = 1,
    Diamond = 2,
    Triangle = 3,
    X = 4,
    Star = 5,
    Dot = 6,
    Circle = 7,
    Plus = 8
}

public enum LineStyle : byte
{
    Solid = 0,
    Dash = 1,
    Dot = 2,
    DashDot = 3,
    DashDotDot = 4,
    None = 5
}

public readonly record struct ChartColor(byte R, byte G, byte B)
{
    public static ChartColor FromRgb(byte r, byte g, byte b) => new(r, g, b);

    public static readonly ChartColor Red = new(255, 0, 0);
    public static readonly ChartColor Green = new(0, 255, 0);
    public static readonly ChartColor Blue = new(0, 0, 255);
    public static readonly ChartColor Yellow = new(255, 255, 0);
    public static readonly ChartColor Cyan = new(0, 255, 255);
    public static readonly ChartColor Magenta = new(255, 0, 255);
    public static readonly ChartColor Black = new(0, 0, 0);
    public static readonly ChartColor White = new(255, 255, 255);
    public static readonly ChartColor Gray = new(128, 128, 128);
    public static readonly ChartColor Orange = new(255, 165, 0);
    public static readonly ChartColor Purple = new(128, 0, 128);

    // 深色变体
    public static readonly ChartColor DarkRed = new(139, 0, 0);
    public static readonly ChartColor DarkGreen = new(0, 100, 0);
    public static readonly ChartColor DarkBlue = new(0, 0, 139);
}

public sealed class ChartRange
{
    public string SheetName { get; init; } = "";
    public int FirstRow { get; init; }
    public int FirstCol { get; init; }
    public int LastRow { get; init; }
    public int LastCol { get; init; }

    public bool IsSingleCell => FirstRow == LastRow && FirstCol == LastCol;

    public ChartRange WithSheetName(string name) => new() { SheetName = name, FirstRow = FirstRow, FirstCol = FirstCol, LastRow = LastRow, LastCol = LastCol };
    public ChartRange WithRows(int firstRow, int lastRow) => new() { SheetName = SheetName, FirstRow = firstRow, FirstCol = FirstCol, LastRow = lastRow, LastCol = LastCol };
    public ChartRange WithCols(int firstCol, int lastCol) => new() { SheetName = SheetName, FirstRow = FirstRow, FirstCol = firstCol, LastRow = LastRow, LastCol = lastCol };
    public ChartRange WithCell(int row, int col) => new() { SheetName = SheetName, FirstRow = row, FirstCol = col, LastRow = row, LastCol = col };
}

public sealed class ChartTitle
{
    public string Text { get; set; } = "";
    public ChartPosition Position { get; set; } = new();
}

public sealed class ChartAxis
{
    public AxisType Type { get; set; } = AxisType.Category;
    public AxisPosition Position { get; set; } = AxisPosition.Bottom;
    public string? Title { get; set; }
    public double? MinValue { get; set; }
    public double? MaxValue { get; set; }
    public bool HasMajorGridlines { get; set; } = true;
    public bool HasMinorGridlines { get; set; }
}

public enum AxisType : byte
{
    Category = 0,
    Value = 1,
    Series = 2
}

public enum AxisPosition : byte
{
    Bottom = 0,
    Left = 1,
    Top = 2,
    Right = 3
}

public sealed class ChartLegend
{
    public LegendPosition Position { get; set; } = LegendPosition.Right;
    public bool Show { get; set; } = true;
}

public enum LegendPosition : byte
{
    Right = 0,
    Left = 1,
    Bottom = 2,
    Top = 3,
    Corner = 4
}

public sealed class ChartPlotArea
{
    public int X { get; set; } = 20;
    public int Y { get; set; } = 20;
    public int Width { get; set; } = 360;
    public int Height { get; set; } = 240;
    public bool VaryColors { get; set; }
}
