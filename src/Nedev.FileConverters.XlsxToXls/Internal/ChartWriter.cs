using System.Buffers;
using System.Buffers.Binary;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// BIFF8图表记录写入器 - 使用ArrayPool减少内存分配
/// </summary>
internal ref struct ChartWriter
{
    private Span<byte> _buffer;
    private int _position;
    private byte[]? _pooledBuffer;

    public ChartWriter(Span<byte> buffer)
    {
        _buffer = buffer;
        _position = 0;
        _pooledBuffer = null;
    }

    /// <summary>
    /// 使用ArrayPool创建ChartWriter，自动管理缓冲区
    /// </summary>
    public static ChartWriter CreatePooled(out byte[] pooledBuffer, int minSize = 65536)
    {
        pooledBuffer = ArrayPool<byte>.Shared.Rent(minSize);
        return new ChartWriter(pooledBuffer.AsSpan())
        {
            _pooledBuffer = pooledBuffer
        };
    }

    /// <summary>
    /// 释放ArrayPool缓冲区（如果使用了CreatePooled）
    /// </summary>
    public void Dispose()
    {
        if (_pooledBuffer != null)
        {
            ArrayPool<byte>.Shared.Return(_pooledBuffer);
            _pooledBuffer = null;
        }
    }

    public int Position => _position;

    /// <summary>
    /// 写入完整的图表流
    /// </summary>
    public int WriteChartStream(ChartData chart, int sheetIndex)
    {
        WriteBofChart();
        WriteChartTypeRecord(chart.Type);

        // 写入图表标题
        if (!string.IsNullOrEmpty(chart.Title?.Text))
        {
            WriteChartTitle(chart.Title.Text);
        }

        // 写入图例
        if (chart.Legend?.Show == true)
        {
            WriteLegend(chart.Legend.Position);
        }

        // 写入绘图区
        WritePlotArea(chart);

        // 写入数据系列
        for (var i = 0; i < chart.Series.Count; i++)
        {
            WriteSeries(chart.Series[i], i);
        }

        // 写入坐标轴
        if (chart.CategoryAxis != null)
        {
            WriteAxis(chart.CategoryAxis);
        }
        if (chart.ValueAxis != null)
        {
            WriteAxis(chart.ValueAxis);
        }

        // 写入系列到轴的关联
        WriteAxisLink();

        WriteEof();
        return _position;
    }

    private void WriteBofChart()
    {
        WriteRecordHeader(0x0809, 16);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0600);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0020);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0C0A);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x07CC);
        _position += 2;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0x00000001);
        _position += 4;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0006);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0000);
        _position += 2;
    }

    private void WriteChartTypeRecord(ChartType type)
    {
        // CHARTTYPE记录 (0x1000系列)
        var recordType = type switch
        {
            ChartType.Area => 0x101A,
            ChartType.Bar => 0x1017,
            ChartType.Line => 0x1018,
            ChartType.Pie => 0x1019,
            ChartType.Scatter => 0x101B,
            ChartType.Radar => 0x103C,
            ChartType.Column => 0x1017, // 柱状图使用Bar记录，通过标志位区分
            ChartType.Doughnut => 0x102C,
            _ => 0x1017
        };

        WriteRecordHeader((ushort)recordType, 6);

        // 图表类型标志
        var flags = type == ChartType.Column ? 0x0001 : 0x0000;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)flags);
        _position += 2;

        // 预留字段
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;
    }

    private void WriteChartTitle(string title)
    {
        // CHARTTITLE记录 (0x102D)
        var bytes = Encoding.Unicode.GetBytes(title);
        var recLen = 4 + bytes.Length;

        WriteRecordHeader(0x102D, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)title.Length);
        _position += 2;
        _buffer[_position++] = 1; // Unicode标志
        bytes.CopyTo(_buffer.Slice(_position));
        _position += bytes.Length;
    }

    private void WriteLegend(LegendPosition position)
    {
        // LEGEND记录 (0x1041)
        WriteRecordHeader(0x1041, 12);

        // 位置
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)position);
        _position += 4;

        // 标志位
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 4;

        // 预留
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;
    }

    private void WritePlotArea(ChartData chart)
    {
        // PLOTAREA记录 (0x1035)
        WriteRecordHeader(0x1035, 16);

        // 绘图区位置和大小（以1/4000为单位）
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)chart.PlotArea.X);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)chart.PlotArea.Y);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)chart.PlotArea.Width);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)chart.PlotArea.Height);
        _position += 4;
    }

    private void WriteSeries(ChartSeries series, int seriesIndex)
    {
        // SERIES记录 (0x1003)
        WriteRecordHeader(0x1003, 8);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)seriesIndex);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(series.CategoryIndex));
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(series.ValueIndex));
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(series.BubbleIndex));
        _position += 2;

        // 系列名称
        if (!string.IsNullOrEmpty(series.Name))
        {
            WriteSeriesName(series.Name);
        }

        // 类别数据
        if (series.Categories != null)
        {
            WriteCategoryRange(series.Categories);
        }

        // 数值数据
        if (series.Values != null)
        {
            WriteValueRange(series.Values);
        }

        // 系列结束标记
        WriteRecordHeader(0x1004, 0);
    }

    private void WriteSeriesName(string name)
    {
        // SERIESTEXT记录 (0x100D)
        var bytes = Encoding.Unicode.GetBytes(name);
        var recLen = 4 + bytes.Length;

        WriteRecordHeader(0x100D, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)name.Length);
        _position += 2;
        _buffer[_position++] = 1; // Unicode
        bytes.CopyTo(_buffer.Slice(_position));
        _position += bytes.Length;
    }

    private void WriteCategoryRange(ChartRange range)
    {
        // CATEGORY记录 (0x1012)
        WriteRecordHeader(0x1012, 12);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstCol);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastCol);
        _position += 2;

        // 引用类型标志
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 2;

        // 工作表索引
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteValueRange(ChartRange range)
    {
        // VALUES记录 (0x1013)
        WriteRecordHeader(0x1013, 12);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstCol);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastCol);
        _position += 2;

        // 引用类型标志
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 2;

        // 工作表索引
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteAxis(ChartAxis axis)
    {
        // AXIS记录 (0x101D)
        WriteRecordHeader(0x101D, 18);

        // 轴类型
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)axis.Type);
        _position += 2;

        // 轴位置
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)axis.Position);
        _position += 2;

        // 标志位
        var flags = 0u;
        if (axis.HasMajorGridlines) flags |= 0x0001;
        if (axis.HasMinorGridlines) flags |= 0x0002;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 最小值/最大值 (如果是数值轴)
        if (axis.MinValue.HasValue)
        {
            BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), axis.MinValue.Value);
        }
        else
        {
            BinaryPrimitives.WriteUInt64LittleEndian(_buffer.Slice(_position), 0xFFFFFFFFFFFFFFFF);
        }
        _position += 8;

        if (axis.MaxValue.HasValue)
        {
            BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), axis.MaxValue.Value);
        }
        else
        {
            BinaryPrimitives.WriteUInt64LittleEndian(_buffer.Slice(_position), 0xFFFFFFFFFFFFFFFF);
        }
        _position += 8;

        // 轴标题
        if (!string.IsNullOrEmpty(axis.Title))
        {
            WriteAxisTitle(axis.Title);
        }

        // 轴结束标记
        WriteRecordHeader(0x101E, 0);
    }

    private void WriteAxisTitle(string title)
    {
        // AXISTITLE记录 (0x102E)
        var bytes = Encoding.Unicode.GetBytes(title);
        var recLen = 4 + bytes.Length;

        WriteRecordHeader(0x102E, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)title.Length);
        _position += 2;
        _buffer[_position++] = 1; // Unicode
        bytes.CopyTo(_buffer.Slice(_position));
        _position += bytes.Length;
    }

    private void WriteAxisLink()
    {
        // AXISLINK记录 (0x1026)
        WriteRecordHeader(0x1026, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteEof()
    {
        WriteRecordHeader(0x000A, 0);
    }

    private void WriteRecordHeader(ushort type, int length)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), type);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)length);
        _position += 2;
    }
}
