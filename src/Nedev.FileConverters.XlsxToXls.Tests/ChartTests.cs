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
    }
}
