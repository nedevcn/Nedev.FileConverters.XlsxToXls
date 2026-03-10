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
    }
}
