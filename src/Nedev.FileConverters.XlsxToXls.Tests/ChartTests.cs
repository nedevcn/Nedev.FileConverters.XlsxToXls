using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using Nedev.FileConverters.XlsxToXls;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    public class ChartTests
    {
        [Fact]
        public void ChartDataModel_Creation()
        {
            var chart = new ChartData
            {
                Name = "TestChart",
                Type = ChartType.Column,
                Title = new ChartTitle { Text = "Sales Chart" },
                Position = new ChartPosition { X = 100, Y = 100, Width = 400, Height = 300 }
            };

            Assert.Equal("TestChart", chart.Name);
            Assert.Equal(ChartType.Column, chart.Type);
            Assert.Equal("Sales Chart", chart.Title?.Text);
            Assert.Equal(100, chart.Position.X);
            Assert.Equal(400, chart.Position.Width);
        }

        [Fact]
        public void ChartSeries_Creation()
        {
            var series = new ChartSeries
            {
                Name = "Q1 Sales",
                Categories = new ChartRange
                {
                    SheetName = "Sheet1",
                    FirstRow = 0,
                    FirstCol = 0,
                    LastRow = 0,
                    LastCol = 3
                },
                Values = new ChartRange
                {
                    SheetName = "Sheet1",
                    FirstRow = 1,
                    FirstCol = 0,
                    LastRow = 1,
                    LastCol = 3
                }
            };

            Assert.Equal("Q1 Sales", series.Name);
            Assert.NotNull(series.Categories);
            Assert.NotNull(series.Values);
            Assert.Equal("Sheet1", series.Categories.SheetName);
        }

        [Fact]
        public void ChartWriter_WritesBofChart()
        {
            var buffer = new byte[4096];
            var writer = new ChartWriter(buffer.AsSpan());
            var chart = new ChartData
            {
                Name = "Test",
                Type = ChartType.Column,
                Series = new List<ChartSeries>()
            };

            var bytesWritten = writer.WriteChartStream(chart, 0);

            Assert.True(bytesWritten > 0);
            // BOF记录类型应该是0x0809
            Assert.Equal(0x09, buffer[0]);
            Assert.Equal(0x08, buffer[1]);
        }

        [Fact]
        public void ChartWriter_WritesChartType()
        {
            var buffer = new byte[4096];
            var writer = new ChartWriter(buffer.AsSpan());
            var chart = new ChartData
            {
                Name = "Test",
                Type = ChartType.Pie,
                Series = new List<ChartSeries>()
            };

            var bytesWritten = writer.WriteChartStream(chart, 0);
            Assert.True(bytesWritten > 0);
        }

        [Fact]
        public void ChartWriter_WritesTitle()
        {
            var buffer = new byte[4096];
            var writer = new ChartWriter(buffer.AsSpan());
            var chart = new ChartData
            {
                Name = "Test",
                Type = ChartType.Column,
                Title = new ChartTitle { Text = "My Chart" },
                Series = new List<ChartSeries>()
            };

            var bytesWritten = writer.WriteChartStream(chart, 0);
            Assert.True(bytesWritten > 0);
        }

        [Fact]
        public void ChartWriter_WritesSeries()
        {
            var buffer = new byte[8192];
            var writer = new ChartWriter(buffer.AsSpan());
            var chart = new ChartData
            {
                Name = "Test",
                Type = ChartType.Column,
                Series = new List<ChartSeries>
                {
                    new()
                    {
                        Name = "Series1",
                        Categories = new ChartRange
                        {
                            FirstRow = 0, FirstCol = 0, LastRow = 0, LastCol = 2
                        },
                        Values = new ChartRange
                        {
                            FirstRow = 1, FirstCol = 0, LastRow = 1, LastCol = 2
                        }
                    }
                }
            };

            var bytesWritten = writer.WriteChartStream(chart, 0);
            Assert.True(bytesWritten > 0);
        }

        [Fact]
        public void ChartWriter_WritesLegend()
        {
            var buffer = new byte[4096];
            var writer = new ChartWriter(buffer.AsSpan());
            var chart = new ChartData
            {
                Name = "Test",
                Type = ChartType.Column,
                Legend = new ChartLegend { Show = true, Position = LegendPosition.Right },
                Series = new List<ChartSeries>()
            };

            var bytesWritten = writer.WriteChartStream(chart, 0);
            Assert.True(bytesWritten > 0);
        }

        [Fact]
        public void ChartWriter_WritesAxes()
        {
            var buffer = new byte[4096];
            var writer = new ChartWriter(buffer.AsSpan());
            var chart = new ChartData
            {
                Name = "Test",
                Type = ChartType.Column,
                CategoryAxis = new ChartAxis
                {
                    Type = AxisType.Category,
                    Position = AxisPosition.Bottom,
                    HasMajorGridlines = false
                },
                ValueAxis = new ChartAxis
                {
                    Type = AxisType.Value,
                    Position = AxisPosition.Left,
                    HasMajorGridlines = true,
                    MinValue = 0,
                    MaxValue = 100
                },
                Series = new List<ChartSeries>()
            };

            var bytesWritten = writer.WriteChartStream(chart, 0);
            Assert.True(bytesWritten > 0);
        }

        [Fact]
        public void ChartRange_IsSingleCell_DetectsCorrectly()
        {
            var singleCell = new ChartRange
            {
                FirstRow = 5,
                FirstCol = 3,
                LastRow = 5,
                LastCol = 3
            };

            var multiCell = new ChartRange
            {
                FirstRow = 0,
                FirstCol = 0,
                LastRow = 5,
                LastCol = 3
            };

            Assert.True(singleCell.IsSingleCell);
            Assert.False(multiCell.IsSingleCell);
        }

        [Theory]
        [InlineData(ChartType.Area)]
        [InlineData(ChartType.Bar)]
        [InlineData(ChartType.Line)]
        [InlineData(ChartType.Pie)]
        [InlineData(ChartType.Scatter)]
        [InlineData(ChartType.Radar)]
        [InlineData(ChartType.Column)]
        [InlineData(ChartType.Doughnut)]
        public void ChartWriter_SupportsAllChartTypes(ChartType type)
        {
            var buffer = new byte[4096];
            var writer = new ChartWriter(buffer.AsSpan());
            var chart = new ChartData
            {
                Name = "Test",
                Type = type,
                Series = new List<ChartSeries>()
            };

            var bytesWritten = writer.WriteChartStream(chart, 0);
            Assert.True(bytesWritten > 0);
        }

        [Fact]
        public void BiffWriter_WritesObjChart()
        {
            var buffer = new byte[4096];
            var writer = new BiffWriter(buffer.AsSpan());
            var chart = new ChartData
            {
                Name = "TestChart",
                Position = new ChartPosition { X = 100, Y = 200, Width = 400, Height = 300 }
            };

            writer.WriteObjChart(1, chart);

            Assert.True(writer.Position > 0);
        }

        [Fact]
        public void ChartData_SupportsMultipleSeries()
        {
            var chart = new ChartData
            {
                Name = "MultiSeriesChart",
                Type = ChartType.Line,
                Series = new List<ChartSeries>
                {
                    new() { Name = "Series1", SeriesIndex = 0 },
                    new() { Name = "Series2", SeriesIndex = 1 },
                    new() { Name = "Series3", SeriesIndex = 2 }
                }
            };

            Assert.Equal(3, chart.Series.Count);
            Assert.Equal("Series2", chart.Series[1].Name);
        }

        [Fact]
        public void ChartAxis_DefaultValues()
        {
            var axis = new ChartAxis
            {
                Type = AxisType.Value,
                Position = AxisPosition.Left
            };

            Assert.Equal(AxisType.Value, axis.Type);
            Assert.Equal(AxisPosition.Left, axis.Position);
            Assert.True(axis.HasMajorGridlines);
            Assert.False(axis.HasMinorGridlines);
            Assert.Null(axis.MinValue);
            Assert.Null(axis.MaxValue);
        }

        [Fact]
        public void ChartLegend_DefaultValues()
        {
            var legend = new ChartLegend();

            Assert.Equal(LegendPosition.Right, legend.Position);
            Assert.True(legend.Show);
        }

        [Fact]
        public void ChartPlotArea_DefaultValues()
        {
            var plotArea = new ChartPlotArea();

            Assert.Equal(20, plotArea.X);
            Assert.Equal(20, plotArea.Y);
            Assert.Equal(360, plotArea.Width);
            Assert.Equal(240, plotArea.Height);
            Assert.False(plotArea.VaryColors);
        }

        // 新增测试：数据标签功能
        [Fact]
        public void DataLabels_DefaultValues()
        {
            var labels = new DataLabels();

            Assert.True(labels.Show);
            Assert.True(labels.ShowValue);
            Assert.False(labels.ShowCategory);
            Assert.False(labels.ShowPercentage);
            Assert.False(labels.ShowSeriesName);
            Assert.Equal(DataLabelPosition.OutsideEnd, labels.Position);
        }

        [Fact]
        public void ChartSeries_WithDataLabels()
        {
            var series = new ChartSeries
            {
                Name = "TestSeries",
                DataLabels = new DataLabels
                {
                    Show = true,
                    ShowValue = true,
                    ShowPercentage = true,
                    Position = DataLabelPosition.InsideEnd
                }
            };

            Assert.NotNull(series.DataLabels);
            Assert.True(series.DataLabels.ShowPercentage);
            Assert.Equal(DataLabelPosition.InsideEnd, series.DataLabels.Position);
        }

        // 新增测试：图表颜色
        [Fact]
        public void ChartColor_PredefinedColors()
        {
            Assert.Equal((255, 0, 0), (ChartColor.Red.R, ChartColor.Red.G, ChartColor.Red.B));
            Assert.Equal((0, 255, 0), (ChartColor.Green.R, ChartColor.Green.G, ChartColor.Green.B));
            Assert.Equal((0, 0, 255), (ChartColor.Blue.R, ChartColor.Blue.G, ChartColor.Blue.B));
            Assert.Equal((0, 0, 0), (ChartColor.Black.R, ChartColor.Black.G, ChartColor.Black.B));
            Assert.Equal((255, 255, 255), (ChartColor.White.R, ChartColor.White.G, ChartColor.White.B));
        }

        [Fact]
        public void ChartColor_FromRgb()
        {
            var color = ChartColor.FromRgb(128, 64, 32);
            Assert.Equal(128, color.R);
            Assert.Equal(64, color.G);
            Assert.Equal(32, color.B);
        }

        [Fact]
        public void ChartSeries_WithColors()
        {
            var series = new ChartSeries
            {
                Name = "ColoredSeries",
                FillColor = ChartColor.Red,
                BorderColor = ChartColor.Black,
                LineStyle = LineStyle.Solid,
                MarkerStyle = MarkerStyle.Circle
            };

            Assert.Equal(ChartColor.Red, series.FillColor);
            Assert.Equal(ChartColor.Black, series.BorderColor);
            Assert.Equal(LineStyle.Solid, series.LineStyle);
            Assert.Equal(MarkerStyle.Circle, series.MarkerStyle);
        }

        // 新增测试：ChartWriter 使用 ArrayPool
        [Fact]
        public void ChartWriter_CreatePooled_WritesChart()
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var chart = new ChartData
                {
                    Name = "PooledChart",
                    Type = ChartType.Column,
                    Series = new List<ChartSeries>()
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
                Assert.True(buffer.Length >= 8192);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：图表范围 With 方法
        [Fact]
        public void ChartRange_WithMethods()
        {
            var original = new ChartRange
            {
                SheetName = "Sheet1",
                FirstRow = 1,
                FirstCol = 2,
                LastRow = 10,
                LastCol = 5
            };

            var withSheet = original.WithSheetName("Sheet2");
            Assert.Equal("Sheet2", withSheet.SheetName);
            Assert.Equal(original.FirstRow, withSheet.FirstRow);

            var withRows = original.WithRows(5, 15);
            Assert.Equal(5, withRows.FirstRow);
            Assert.Equal(15, withRows.LastRow);

            var withCols = original.WithCols(3, 8);
            Assert.Equal(3, withCols.FirstCol);
            Assert.Equal(8, withCols.LastCol);

            var withCell = original.WithCell(20, 30);
            Assert.Equal(20, withCell.FirstRow);
            Assert.Equal(20, withCell.LastRow);
            Assert.Equal(30, withCell.FirstCol);
            Assert.Equal(30, withCell.LastCol);
        }

        // 新增测试：枚举值验证
        [Theory]
        [InlineData(AxisType.Category, 0)]
        [InlineData(AxisType.Value, 1)]
        [InlineData(AxisType.Series, 2)]
        public void AxisType_HasCorrectValues(AxisType type, byte expected)
        {
            Assert.Equal(expected, (byte)type);
        }

        [Theory]
        [InlineData(AxisPosition.Bottom, 0)]
        [InlineData(AxisPosition.Left, 1)]
        [InlineData(AxisPosition.Top, 2)]
        [InlineData(AxisPosition.Right, 3)]
        public void AxisPosition_HasCorrectValues(AxisPosition position, byte expected)
        {
            Assert.Equal(expected, (byte)position);
        }

        [Theory]
        [InlineData(LegendPosition.Right, 0)]
        [InlineData(LegendPosition.Left, 1)]
        [InlineData(LegendPosition.Bottom, 2)]
        [InlineData(LegendPosition.Top, 3)]
        [InlineData(LegendPosition.Corner, 4)]
        public void LegendPosition_HasCorrectValues(LegendPosition position, byte expected)
        {
            Assert.Equal(expected, (byte)position);
        }

        // 新增测试：复杂图表配置
        [Fact]
        public void ChartData_ComplexConfiguration()
        {
            var chart = new ChartData
            {
                Name = "ComplexChart",
                Type = ChartType.Line,
                Title = new ChartTitle
                {
                    Text = "Sales Report",
                    Position = new ChartPosition { X = 100, Y = 10 }
                },
                Legend = new ChartLegend
                {
                    Show = true,
                    Position = LegendPosition.Bottom
                },
                CategoryAxis = new ChartAxis
                {
                    Type = AxisType.Category,
                    Position = AxisPosition.Bottom,
                    Title = "Months",
                    HasMajorGridlines = false
                },
                ValueAxis = new ChartAxis
                {
                    Type = AxisType.Value,
                    Position = AxisPosition.Left,
                    Title = "Revenue",
                    MinValue = 0,
                    MaxValue = 100000,
                    HasMajorGridlines = true
                },
                Series = new List<ChartSeries>
                {
                    new()
                    {
                        Name = "Q1 Sales",
                        SeriesIndex = 0,
                        FillColor = ChartColor.Blue,
                        LineStyle = LineStyle.Solid,
                        MarkerStyle = MarkerStyle.Diamond,
                        DataLabels = new DataLabels
                        {
                            Show = true,
                            ShowValue = true,
                            Position = DataLabelPosition.Above
                        }
                    },
                    new()
                    {
                        Name = "Q2 Sales",
                        SeriesIndex = 1,
                        FillColor = ChartColor.Green,
                        LineStyle = LineStyle.Dash,
                        MarkerStyle = MarkerStyle.Square
                    }
                }
            };

            Assert.Equal("ComplexChart", chart.Name);
            Assert.Equal(ChartType.Line, chart.Type);
            Assert.Equal("Sales Report", chart.Title?.Text);
            Assert.Equal(LegendPosition.Bottom, chart.Legend?.Position);
            Assert.Equal(2, chart.Series.Count);
            Assert.Equal(LineStyle.Dash, chart.Series[1].LineStyle);
            Assert.Equal(DataLabelPosition.Above, chart.Series[0].DataLabels?.Position);
        }

        // 新增测试：边界情况
        [Fact]
        public void ChartData_EmptySeriesList()
        {
            var chart = new ChartData
            {
                Name = "EmptyChart",
                Type = ChartType.Pie,
                Series = new List<ChartSeries>()
            };

            Assert.Empty(chart.Series);
        }

        [Fact]
        public void ChartPosition_ZeroDimensions()
        {
            var position = new ChartPosition
            {
                X = 0,
                Y = 0,
                Width = 0,
                Height = 0
            };

            Assert.Equal(0, position.X);
            Assert.Equal(0, position.Y);
            Assert.Equal(0, position.Width);
            Assert.Equal(0, position.Height);
        }

        [Fact]
        public void ChartTitle_EmptyText()
        {
            var title = new ChartTitle
            {
                Text = "",
                Position = new ChartPosition()
            };

            Assert.Equal("", title.Text);
        }

        // 新增测试：验证 ChartWriter 写入数据标签
        [Fact]
        public void ChartWriter_WritesDataLabels()
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var chart = new ChartData
                {
                    Name = "ChartWithLabels",
                    Type = ChartType.Column,
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "Series1",
                            SeriesIndex = 0,
                            DataLabels = new DataLabels
                            {
                                Show = true,
                                ShowValue = true,
                                ShowPercentage = true,
                                Position = DataLabelPosition.InsideEnd
                            }
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：验证 ChartWriter 写入系列样式
        [Fact]
        public void ChartWriter_WritesSeriesStyle()
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var chart = new ChartData
                {
                    Name = "StyledChart",
                    Type = ChartType.Line,
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "StyledSeries",
                            SeriesIndex = 0,
                            LineStyle = LineStyle.Dash,
                            MarkerStyle = MarkerStyle.Diamond,
                            FillColor = ChartColor.Red,
                            BorderColor = ChartColor.Black
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：验证 ChartWriter 写入系列颜色
        [Theory]
        [InlineData(255, 0, 0)]    // Red
        [InlineData(0, 255, 0)]    // Green
        [InlineData(0, 0, 255)]    // Blue
        [InlineData(255, 255, 0)]  // Yellow
        [InlineData(128, 64, 32)]  // Custom
        public void ChartWriter_WritesSeriesColor(byte r, byte g, byte b)
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var chart = new ChartData
                {
                    Name = "ColoredChart",
                    Type = ChartType.Column,
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "ColoredSeries",
                            SeriesIndex = 0,
                            FillColor = new ChartColor(r, g, b)
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：验证所有线条样式
        [Theory]
        [InlineData(LineStyle.Solid)]
        [InlineData(LineStyle.Dash)]
        [InlineData(LineStyle.Dot)]
        [InlineData(LineStyle.DashDot)]
        [InlineData(LineStyle.DashDotDot)]
        [InlineData(LineStyle.None)]
        public void ChartWriter_WritesAllLineStyles(LineStyle style)
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var chart = new ChartData
                {
                    Name = "LineStyleChart",
                    Type = ChartType.Line,
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "LineSeries",
                            SeriesIndex = 0,
                            LineStyle = style
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：验证所有标记样式
        [Theory]
        [InlineData(MarkerStyle.None)]
        [InlineData(MarkerStyle.Square)]
        [InlineData(MarkerStyle.Diamond)]
        [InlineData(MarkerStyle.Triangle)]
        [InlineData(MarkerStyle.X)]
        [InlineData(MarkerStyle.Star)]
        [InlineData(MarkerStyle.Dot)]
        [InlineData(MarkerStyle.Circle)]
        [InlineData(MarkerStyle.Plus)]
        public void ChartWriter_WritesAllMarkerStyles(MarkerStyle style)
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var chart = new ChartData
                {
                    Name = "MarkerStyleChart",
                    Type = ChartType.Line,
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "MarkerSeries",
                            SeriesIndex = 0,
                            MarkerStyle = style
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：验证所有数据标签位置
        [Theory]
        [InlineData(DataLabelPosition.Center)]
        [InlineData(DataLabelPosition.InsideEnd)]
        [InlineData(DataLabelPosition.OutsideEnd)]
        [InlineData(DataLabelPosition.BestFit)]
        [InlineData(DataLabelPosition.Left)]
        [InlineData(DataLabelPosition.Right)]
        [InlineData(DataLabelPosition.Above)]
        [InlineData(DataLabelPosition.Below)]
        public void ChartWriter_WritesAllDataLabelPositions(DataLabelPosition position)
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var chart = new ChartData
                {
                    Name = "LabelPositionChart",
                    Type = ChartType.Column,
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "LabelSeries",
                            SeriesIndex = 0,
                            DataLabels = new DataLabels
                            {
                                Show = true,
                                ShowValue = true,
                                Position = position
                            }
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：复杂图表 - 多系列带完整样式
        [Fact]
        public void ChartWriter_ComplexMultiSeriesChart()
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 16384);
            try
            {
                var chart = new ChartData
                {
                    Name = "ComplexMultiSeries",
                    Type = ChartType.Line,
                    Title = new ChartTitle { Text = "Sales Performance" },
                    Legend = new ChartLegend { Show = true, Position = LegendPosition.Bottom },
                    CategoryAxis = new ChartAxis
                    {
                        Type = AxisType.Category,
                        Position = AxisPosition.Bottom,
                        Title = "Months",
                        HasMajorGridlines = false
                    },
                    ValueAxis = new ChartAxis
                    {
                        Type = AxisType.Value,
                        Position = AxisPosition.Left,
                        Title = "Revenue ($)",
                        MinValue = 0,
                        MaxValue = 100000,
                        HasMajorGridlines = true
                    },
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "Product A",
                            SeriesIndex = 0,
                            FillColor = ChartColor.Blue,
                            BorderColor = ChartColor.DarkBlue,
                            LineStyle = LineStyle.Solid,
                            MarkerStyle = MarkerStyle.Circle,
                            DataLabels = new DataLabels
                            {
                                Show = true,
                                ShowValue = true,
                                Position = DataLabelPosition.Above
                            }
                        },
                        new()
                        {
                            Name = "Product B",
                            SeriesIndex = 1,
                            FillColor = ChartColor.Green,
                            BorderColor = ChartColor.DarkGreen,
                            LineStyle = LineStyle.Dash,
                            MarkerStyle = MarkerStyle.Square,
                            DataLabels = new DataLabels
                            {
                                Show = true,
                                ShowValue = true,
                                Position = DataLabelPosition.Below
                            }
                        },
                        new()
                        {
                            Name = "Product C",
                            SeriesIndex = 2,
                            FillColor = ChartColor.Red,
                            BorderColor = ChartColor.DarkRed,
                            LineStyle = LineStyle.Dot,
                            MarkerStyle = MarkerStyle.Diamond,
                            DataLabels = new DataLabels
                            {
                                Show = true,
                                ShowPercentage = true,
                                Position = DataLabelPosition.Center
                            }
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 100); // 确保写入了足够多的数据
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：ChartColor 相等性和哈希码
        [Fact]
        public void ChartColor_Equality()
        {
            var color1 = new ChartColor(255, 128, 64);
            var color2 = new ChartColor(255, 128, 64);
            var color3 = new ChartColor(255, 128, 65);

            Assert.Equal(color1, color2);
            Assert.True(color1 == color2);
            Assert.False(color1 == color3);
            Assert.Equal(color1.GetHashCode(), color2.GetHashCode());
        }

        // 新增测试：ChartColor ToString
        [Fact]
        public void ChartColor_ToString()
        {
            var color = new ChartColor(255, 128, 64);
            var str = color.ToString();
            Assert.Contains("255", str);
            Assert.Contains("128", str);
            Assert.Contains("64", str);
        }

        // 新增测试：数据点级别支持
        [Fact]
        public void ChartDataPoint_DefaultValues()
        {
            var point = new ChartDataPoint();
            Assert.Equal(0, point.Index);
            Assert.Null(point.FillColor);
            Assert.Null(point.BorderColor);
            Assert.Null(point.DataLabels);
            Assert.Null(point.Explosion);
        }

        [Fact]
        public void ChartSeries_WithDataPoints()
        {
            var series = new ChartSeries
            {
                Name = "SeriesWithPoints",
                DataPoints = new List<ChartDataPoint>
                {
                    new() { Index = 0, FillColor = ChartColor.Red },
                    new() { Index = 1, FillColor = ChartColor.Green },
                    new() { Index = 2, FillColor = ChartColor.Blue, Explosion = true }
                }
            };

            Assert.Equal(3, series.DataPoints.Count);
            Assert.Equal(ChartColor.Red, series.DataPoints[0].FillColor);
            Assert.True(series.DataPoints[2].Explosion);
        }

        // 新增测试：趋势线
        [Fact]
        public void TrendLine_DefaultValues()
        {
            var trendLine = new TrendLine();
            Assert.Equal(TrendLineType.Linear, trendLine.Type);
            Assert.Null(trendLine.Name);
            Assert.False(trendLine.DisplayEquation);
            Assert.False(trendLine.DisplayRSquared);
            Assert.Equal(2, trendLine.Order);
            Assert.Equal(2, trendLine.Period);
        }

        [Theory]
        [InlineData(TrendLineType.Linear)]
        [InlineData(TrendLineType.Exponential)]
        [InlineData(TrendLineType.Logarithmic)]
        [InlineData(TrendLineType.Polynomial)]
        [InlineData(TrendLineType.Power)]
        [InlineData(TrendLineType.MovingAverage)]
        public void TrendLine_AllTypes(TrendLineType type)
        {
            var trendLine = new TrendLine { Type = type };
            Assert.Equal(type, trendLine.Type);
        }

        [Fact]
        public void ChartSeries_WithTrendLines()
        {
            var series = new ChartSeries
            {
                Name = "SeriesWithTrend",
                TrendLines = new List<TrendLine>
                {
                    new()
                    {
                        Type = TrendLineType.Linear,
                        DisplayEquation = true,
                        DisplayRSquared = true,
                        LineColor = ChartColor.Red
                    },
                    new()
                    {
                        Type = TrendLineType.Polynomial,
                        Order = 3,
                        LineColor = ChartColor.Blue
                    }
                }
            };

            Assert.Equal(2, series.TrendLines.Count);
            Assert.True(series.TrendLines[0].DisplayEquation);
            Assert.Equal(3, series.TrendLines[1].Order);
        }

        // 新增测试：误差线
        [Fact]
        public void ErrorBars_DefaultValues()
        {
            var errorBars = new ErrorBars();
            Assert.Equal(ErrorBarType.Both, errorBars.Type);
            Assert.Equal(ErrorBarValueType.FixedValue, errorBars.ValueType);
            Assert.Equal(0, errorBars.Value);
            Assert.True(errorBars.ShowCap);
        }

        [Theory]
        [InlineData(ErrorBarType.Both)]
        [InlineData(ErrorBarType.Plus)]
        [InlineData(ErrorBarType.Minus)]
        public void ErrorBars_AllTypes(ErrorBarType type)
        {
            var errorBars = new ErrorBars { Type = type };
            Assert.Equal(type, errorBars.Type);
        }

        [Theory]
        [InlineData(ErrorBarValueType.FixedValue)]
        [InlineData(ErrorBarValueType.Percentage)]
        [InlineData(ErrorBarValueType.StandardDeviation)]
        [InlineData(ErrorBarValueType.StandardError)]
        [InlineData(ErrorBarValueType.Custom)]
        public void ErrorBars_AllValueTypes(ErrorBarValueType valueType)
        {
            var errorBars = new ErrorBars { ValueType = valueType };
            Assert.Equal(valueType, errorBars.ValueType);
        }

        [Fact]
        public void ChartSeries_WithErrorBars()
        {
            var series = new ChartSeries
            {
                Name = "SeriesWithErrorBars",
                ErrorBars = new ErrorBars
                {
                    Type = ErrorBarType.Both,
                    ValueType = ErrorBarValueType.Percentage,
                    Value = 5.0,
                    ShowCap = true,
                    LineColor = ChartColor.Gray
                }
            };

            Assert.NotNull(series.ErrorBars);
            Assert.Equal(5.0, series.ErrorBars.Value);
            Assert.Equal(ErrorBarValueType.Percentage, series.ErrorBars.ValueType);
        }

        // 新增测试：ChartWriter 写入数据点
        [Fact]
        public void ChartWriter_WritesDataPoints()
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 16384);
            try
            {
                var chart = new ChartData
                {
                    Name = "ChartWithDataPoints",
                    Type = ChartType.Pie,
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "PieSeries",
                            SeriesIndex = 0,
                            DataPoints = new List<ChartDataPoint>
                            {
                                new() { Index = 0, FillColor = ChartColor.Red, Explosion = true },
                                new() { Index = 1, FillColor = ChartColor.Green },
                                new() { Index = 2, FillColor = ChartColor.Blue, DataLabels = new DataLabels { Show = true, ShowValue = true } }
                            }
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：ChartWriter 写入趋势线
        [Fact]
        public void ChartWriter_WritesTrendLines()
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 16384);
            try
            {
                var chart = new ChartData
                {
                    Name = "ChartWithTrendLines",
                    Type = ChartType.Line,
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "SeriesWithTrend",
                            SeriesIndex = 0,
                            TrendLines = new List<TrendLine>
                            {
                                new()
                                {
                                    Type = TrendLineType.Linear,
                                    Name = "Linear Trend",
                                    DisplayEquation = true,
                                    DisplayRSquared = true,
                                    LineColor = ChartColor.Red,
                                    LineStyle = LineStyle.Dash
                                }
                            }
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：ChartWriter 写入误差线
        [Fact]
        public void ChartWriter_WritesErrorBars()
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 16384);
            try
            {
                var chart = new ChartData
                {
                    Name = "ChartWithErrorBars",
                    Type = ChartType.Column,
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "SeriesWithErrors",
                            SeriesIndex = 0,
                            ErrorBars = new ErrorBars
                            {
                                Type = ErrorBarType.Both,
                                ValueType = ErrorBarValueType.StandardDeviation,
                                Value = 1.0,
                                ShowCap = true,
                                LineColor = ChartColor.Gray
                            }
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：所有趋势线类型写入
        [Theory]
        [InlineData(TrendLineType.Linear)]
        [InlineData(TrendLineType.Exponential)]
        [InlineData(TrendLineType.Logarithmic)]
        [InlineData(TrendLineType.Polynomial)]
        [InlineData(TrendLineType.Power)]
        [InlineData(TrendLineType.MovingAverage)]
        public void ChartWriter_WritesAllTrendLineTypes(TrendLineType type)
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var chart = new ChartData
                {
                    Name = "TrendLineChart",
                    Type = ChartType.Line,
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "Series",
                            SeriesIndex = 0,
                            TrendLines = new List<TrendLine>
                            {
                                new() { Type = type }
                            }
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：所有误差线类型写入
        [Theory]
        [InlineData(ErrorBarType.Both)]
        [InlineData(ErrorBarType.Plus)]
        [InlineData(ErrorBarType.Minus)]
        public void ChartWriter_WritesAllErrorBarTypes(ErrorBarType type)
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var chart = new ChartData
                {
                    Name = "ErrorBarChart",
                    Type = ChartType.Column,
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "Series",
                            SeriesIndex = 0,
                            ErrorBars = new ErrorBars { Type = type }
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 新增测试：组合图表类型属性
        [Fact]
        public void ChartSeries_SecondaryAxisProperties()
        {
            var series = new ChartSeries
            {
                Name = "SecondarySeries",
                SecondaryChartType = ChartType.Line,
                UseSecondaryAxis = true
            };

            Assert.Equal(ChartType.Line, series.SecondaryChartType);
            Assert.True(series.UseSecondaryAxis);
        }

        // 新增测试：完整功能图表
        [Fact]
        public void ChartWriter_FullFeaturedChart()
        {
            var writer = ChartWriter.CreatePooled(out var buffer, 32768);
            try
            {
                var chart = new ChartData
                {
                    Name = "FullFeaturedChart",
                    Type = ChartType.Column,
                    Title = new ChartTitle { Text = "Complete Chart Example" },
                    Legend = new ChartLegend { Show = true, Position = LegendPosition.Right },
                    CategoryAxis = new ChartAxis
                    {
                        Type = AxisType.Category,
                        Position = AxisPosition.Bottom,
                        Title = "Categories"
                    },
                    ValueAxis = new ChartAxis
                    {
                        Type = AxisType.Value,
                        Position = AxisPosition.Left,
                        Title = "Values",
                        MinValue = 0,
                        MaxValue = 100
                    },
                    Series = new List<ChartSeries>
                    {
                        new()
                        {
                            Name = "Primary Series",
                            SeriesIndex = 0,
                            FillColor = ChartColor.Blue,
                            DataPoints = new List<ChartDataPoint>
                            {
                                new() { Index = 0, FillColor = ChartColor.Red },
                                new() { Index = 1, FillColor = ChartColor.Green },
                                new() { Index = 2, FillColor = ChartColor.Yellow }
                            },
                            TrendLines = new List<TrendLine>
                            {
                                new()
                                {
                                    Type = TrendLineType.Linear,
                                    DisplayEquation = true,
                                    LineColor = ChartColor.DarkBlue
                                }
                            },
                            ErrorBars = new ErrorBars
                            {
                                Type = ErrorBarType.Both,
                                ValueType = ErrorBarValueType.Percentage,
                                Value = 5.0
                            }
                        },
                        new()
                        {
                            Name = "Secondary Series",
                            SeriesIndex = 1,
                            SecondaryChartType = ChartType.Line,
                            UseSecondaryAxis = true,
                            FillColor = ChartColor.Orange,
                            LineStyle = LineStyle.Solid,
                            MarkerStyle = MarkerStyle.Circle
                        }
                    }
                };

                var bytesWritten = writer.WriteChartStream(chart, 0);
                Assert.True(bytesWritten > 200);
            }
            finally
            {
                writer.Dispose();
            }
        }
    }
}
