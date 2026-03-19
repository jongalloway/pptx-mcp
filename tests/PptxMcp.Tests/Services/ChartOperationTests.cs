namespace PptxMcp.Tests.Services;

[Trait("Category", "Unit")]
public class ChartOperationTests : PptxTestBase
{
    // ── GetChartData: happy paths ────────────────────────────────────────────────

    [Fact]
    public void GetChartData_ColumnChart_ReturnsSuccess()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    Name = "Sales Chart",
                    ChartType = "Column",
                    Categories = ["Q1", "Q2", "Q3"],
                    Series = [new TestSeriesDefinition { Name = "Revenue", Values = [100.0, 200.0, 150.0] }]
                }
            ]
        });

        var result = Service.GetChartData(path, slideNumber: 1);

        Assert.True(result.Success);
        Assert.Equal(1, result.SlideNumber);
        Assert.Equal("Sales Chart", result.ChartName);
        Assert.Equal("Column", result.ChartType);
        Assert.Equal(1, result.SeriesCount);
    }

    [Fact]
    public void GetChartData_ColumnChart_ReturnsCategoriesAndValues()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Column",
                    Categories = ["Jan", "Feb", "Mar"],
                    Series = [new TestSeriesDefinition { Name = "Sales", Values = [10.0, 20.0, 15.0] }]
                }
            ]
        });

        var result = Service.GetChartData(path, slideNumber: 1);

        Assert.True(result.Success);
        Assert.Single(result.Series);
        Assert.Equal("Sales", result.Series[0].SeriesName);
        Assert.Equal(["Jan", "Feb", "Mar"], result.Series[0].Categories);
        Assert.Equal([10.0, 20.0, 15.0], result.Series[0].Values);
    }

    [Fact]
    public void GetChartData_BarChart_ReturnsBarChartType()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Bar",
                    Categories = ["A", "B"],
                    Series = [new TestSeriesDefinition { Name = "Series 1", Values = [5.0, 10.0] }]
                }
            ]
        });

        var result = Service.GetChartData(path, slideNumber: 1);

        Assert.True(result.Success);
        Assert.Equal("Bar", result.ChartType);
    }

    [Fact]
    public void GetChartData_LineChart_ReturnsLineChartType()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Line",
                    Categories = ["W1", "W2", "W3"],
                    Series = [new TestSeriesDefinition { Name = "Trend", Values = [1.0, 3.0, 2.0] }]
                }
            ]
        });

        var result = Service.GetChartData(path, slideNumber: 1);

        Assert.True(result.Success);
        Assert.Equal("Line", result.ChartType);
    }

    [Fact]
    public void GetChartData_PieChart_ReturnsPieChartType()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Pie",
                    Categories = ["Slice A", "Slice B", "Slice C"],
                    Series = [new TestSeriesDefinition { Name = "Share", Values = [40.0, 35.0, 25.0] }]
                }
            ]
        });

        var result = Service.GetChartData(path, slideNumber: 1);

        Assert.True(result.Success);
        Assert.Equal("Pie", result.ChartType);
    }

    [Fact]
    public void GetChartData_MultiSeries_ReturnsAllSeries()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Column",
                    Categories = ["Q1", "Q2"],
                    Series =
                    [
                        new TestSeriesDefinition { Name = "Product A", Values = [100.0, 120.0] },
                        new TestSeriesDefinition { Name = "Product B", Values = [80.0, 90.0] }
                    ]
                }
            ]
        });

        var result = Service.GetChartData(path, slideNumber: 1);

        Assert.True(result.Success);
        Assert.Equal(2, result.SeriesCount);
        Assert.Equal(2, result.Series.Length);
        Assert.Equal("Product A", result.Series[0].SeriesName);
        Assert.Equal("Product B", result.Series[1].SeriesName);
    }

    [Fact]
    public void GetChartData_OnlyChart_MatchedByOnlyChart()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts = [new TestChartDefinition { ChartType = "Column", Series = [new TestSeriesDefinition { Values = [1.0] }] }]
        });

        var result = Service.GetChartData(path, slideNumber: 1);

        Assert.Equal("onlyChart", result.MatchedBy);
    }

    [Fact]
    public void GetChartData_ByChartName_MatchedByChartName()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts = [new TestChartDefinition { Name = "My Chart", ChartType = "Column", Series = [new TestSeriesDefinition { Values = [1.0] }] }]
        });

        var result = Service.GetChartData(path, slideNumber: 1, chartName: "My Chart");

        Assert.True(result.Success);
        Assert.Equal("chartName", result.MatchedBy);
    }

    [Fact]
    public void GetChartData_ByChartIndex_MatchedByChartIndex()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts = [new TestChartDefinition { ChartType = "Column", Series = [new TestSeriesDefinition { Values = [1.0] }] }]
        });

        var result = Service.GetChartData(path, slideNumber: 1, chartIndex: 0);

        Assert.True(result.Success);
        Assert.Equal("chartIndex", result.MatchedBy);
    }

    // ── GetChartData: error cases ────────────────────────────────────────────────

    [Fact]
    public void GetChartData_NoCharts_ReturnsFailure()
    {
        var path = CreateMinimalPptx();

        var result = Service.GetChartData(path, slideNumber: 1);

        Assert.False(result.Success);
        Assert.Contains("no charts", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetChartData_ChartNameNotFound_Throws()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts = [new TestChartDefinition { Name = "Real Chart", ChartType = "Column", Series = [new TestSeriesDefinition { Values = [1.0] }] }]
        });

        Assert.Throws<InvalidOperationException>(() =>
            Service.GetChartData(path, slideNumber: 1, chartName: "Nonexistent Chart"));
    }

    [Fact]
    public void GetChartData_ChartIndexOutOfRange_Throws()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts = [new TestChartDefinition { ChartType = "Column", Series = [new TestSeriesDefinition { Values = [1.0] }] }]
        });

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            Service.GetChartData(path, slideNumber: 1, chartIndex: 5));
    }

    // ── UpdateChartData: happy paths ─────────────────────────────────────────────

    [Fact]
    public void UpdateChartData_UpdatesValues_ReadBackConfirmed()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Column",
                    Categories = ["Q1", "Q2", "Q3"],
                    Series = [new TestSeriesDefinition { Name = "Sales", Values = [10.0, 20.0, 30.0] }]
                }
            ]
        });

        var result = Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: null, Values: [50.0, 60.0, 70.0])
        ]);

        Assert.True(result.Success);
        Assert.Equal(1, result.SeriesUpdated);

        // Read back to verify
        var readBack = Service.GetChartData(path, slideNumber: 1);
        Assert.Equal([50.0, 60.0, 70.0], readBack.Series[0].Values);
    }

    [Fact]
    public void UpdateChartData_UpdatesSeriesName_ReadBackConfirmed()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Column",
                    Categories = ["A"],
                    Series = [new TestSeriesDefinition { Name = "Old Name", Values = [1.0] }]
                }
            ]
        });

        Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: "New Name", Categories: null, Values: null)
        ]);

        var readBack = Service.GetChartData(path, slideNumber: 1);
        Assert.Equal("New Name", readBack.Series[0].SeriesName);
    }

    [Fact]
    public void UpdateChartData_UpdatesCategories_ReadBackConfirmed()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Column",
                    Categories = ["Old1", "Old2"],
                    Series = [new TestSeriesDefinition { Name = "S1", Values = [1.0, 2.0] }]
                }
            ]
        });

        Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: ["New1", "New2"], Values: null)
        ]);

        var readBack = Service.GetChartData(path, slideNumber: 1);
        Assert.Equal(["New1", "New2"], readBack.Series[0].Categories);
    }

    [Fact]
    public void UpdateChartData_ChangeValueCount_ReadBackConfirmed()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Column",
                    Categories = ["A", "B", "C"],
                    Series = [new TestSeriesDefinition { Name = "S1", Values = [1.0, 2.0, 3.0] }]
                }
            ]
        });

        Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: ["X", "Y"], Values: [9.0, 8.0])
        ]);

        var readBack = Service.GetChartData(path, slideNumber: 1);
        Assert.Equal(2, readBack.Series[0].Values.Length);
        Assert.Equal([9.0, 8.0], readBack.Series[0].Values);
    }

    [Fact]
    public void UpdateChartData_SkipsOutOfRangeSeries_ReturnsPartialSuccess()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Column",
                    Series = [new TestSeriesDefinition { Name = "S1", Values = [1.0] }]
                }
            ]
        });

        var result = Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: null, Values: [99.0]),
            new ChartSeriesUpdate(SeriesIndex: 5, SeriesName: null, Categories: null, Values: [0.0]) // out of range
        ]);

        Assert.True(result.Success);
        Assert.Equal(1, result.SeriesUpdated); // only the valid one counted
    }

    [Fact]
    public void UpdateChartData_PieChart_UpdatesValuesAndCategories()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Pie",
                    Categories = ["A", "B"],
                    Series = [new TestSeriesDefinition { Name = "Share", Values = [50.0, 50.0] }]
                }
            ]
        });

        Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: ["X", "Y", "Z"], Values: [30.0, 40.0, 30.0])
        ]);

        var readBack = Service.GetChartData(path, slideNumber: 1);
        Assert.Equal(3, readBack.Series[0].Values.Length);
        Assert.Equal([30.0, 40.0, 30.0], readBack.Series[0].Values);
    }

    [Fact]
    public void UpdateChartData_LineChart_UpdatesValues()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Line",
                    Categories = ["Jan", "Feb", "Mar"],
                    Series = [new TestSeriesDefinition { Name = "Trend", Values = [1.0, 2.0, 3.0] }]
                }
            ]
        });

        Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: null, Values: [10.0, 15.0, 12.0])
        ]);

        var readBack = Service.GetChartData(path, slideNumber: 1);
        Assert.Equal([10.0, 15.0, 12.0], readBack.Series[0].Values);
    }

    // ── UpdateChartData: error cases ─────────────────────────────────────────────

    [Fact]
    public void UpdateChartData_NoCharts_ReturnsFailure()
    {
        var path = CreateMinimalPptx();

        var result = Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: null, Values: [1.0])
        ]);

        Assert.False(result.Success);
        Assert.Contains("no charts", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UpdateChartData_ChartNameNotFound_Throws()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts = [new TestChartDefinition { Name = "Real", ChartType = "Column", Series = [new TestSeriesDefinition { Values = [1.0] }] }]
        });

        Assert.Throws<InvalidOperationException>(() =>
            Service.UpdateChartData(path, slideNumber: 1, [
                new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: null, Values: [2.0])
            ], chartName: "Fake"));
    }

    // ── Preservation of formatting ───────────────────────────────────────────────

    [Fact]
    public void UpdateChartData_PreservesChartType()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Column",
                    Categories = ["A", "B"],
                    Series = [new TestSeriesDefinition { Name = "S1", Values = [1.0, 2.0] }]
                }
            ]
        });

        Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: null, Values: [9.0, 8.0])
        ]);

        var readBack = Service.GetChartData(path, slideNumber: 1);
        Assert.Equal("Column", readBack.ChartType);
    }

    [Fact]
    public void UpdateChartData_MultiSeries_UpdatesOnlyTargetedSeries()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Column",
                    Categories = ["A", "B"],
                    Series =
                    [
                        new TestSeriesDefinition { Name = "S0", Values = [1.0, 2.0] },
                        new TestSeriesDefinition { Name = "S1", Values = [3.0, 4.0] }
                    ]
                }
            ]
        });

        Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 1, SeriesName: null, Categories: null, Values: [30.0, 40.0])
        ]);

        var readBack = Service.GetChartData(path, slideNumber: 1);
        // Series 0 should be untouched
        Assert.Equal([1.0, 2.0], readBack.Series[0].Values);
        // Series 1 should be updated
        Assert.Equal([30.0, 40.0], readBack.Series[1].Values);
    }

    // ── Additional chart types ────────────────────────────────────────────────────

    [Fact]
    public void GetChartData_AreaChart_ReturnsAreaChartType()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Area",
                    Categories = ["Jan", "Feb", "Mar"],
                    Series = [new TestSeriesDefinition { Name = "Sales", Values = [10.0, 20.0, 15.0] }]
                }
            ]
        });

        var result = Service.GetChartData(path, slideNumber: 1);

        Assert.True(result.Success);
        Assert.Equal("Area", result.ChartType);
        Assert.Equal("Sales", result.Series[0].SeriesName);
        Assert.Equal(["Jan", "Feb", "Mar"], result.Series[0].Categories);
        Assert.Equal([10.0, 20.0, 15.0], result.Series[0].Values);
    }

    [Fact]
    public void UpdateChartData_AreaChart_UpdatesValues()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Area",
                    Categories = ["A", "B", "C"],
                    Series = [new TestSeriesDefinition { Name = "S1", Values = [1.0, 2.0, 3.0] }]
                }
            ]
        });

        Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: null, Values: [7.0, 8.0, 9.0])
        ]);

        var readBack = Service.GetChartData(path, slideNumber: 1);
        Assert.Equal("Area", readBack.ChartType);
        Assert.Equal([7.0, 8.0, 9.0], readBack.Series[0].Values);
    }

    [Fact]
    public void GetChartData_DoughnutChart_ReturnsDoughnutChartType()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Doughnut",
                    Categories = ["Slice A", "Slice B"],
                    Series = [new TestSeriesDefinition { Name = "Share", Values = [60.0, 40.0] }]
                }
            ]
        });

        var result = Service.GetChartData(path, slideNumber: 1);

        Assert.True(result.Success);
        Assert.Equal("Doughnut", result.ChartType);
        Assert.Equal([60.0, 40.0], result.Series[0].Values);
    }

    [Fact]
    public void UpdateChartData_DoughnutChart_UpdatesValues()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Doughnut",
                    Categories = ["A", "B"],
                    Series = [new TestSeriesDefinition { Name = "S", Values = [50.0, 50.0] }]
                }
            ]
        });

        Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: null, Values: [70.0, 30.0])
        ]);

        var readBack = Service.GetChartData(path, slideNumber: 1);
        Assert.Equal("Doughnut", readBack.ChartType);
        Assert.Equal([70.0, 30.0], readBack.Series[0].Values);
    }

    [Fact]
    public void GetChartData_ScatterChart_ReturnsScatterChartType()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Scatter",
                    Series = [new TestSeriesDefinition { Name = "Dataset", Values = [10.0, 20.0, 30.0] }]
                }
            ]
        });

        var result = Service.GetChartData(path, slideNumber: 1);

        Assert.True(result.Success);
        Assert.Equal("Scatter", result.ChartType);
        // Scatter returns X values as categories and Y values as values
        Assert.Equal(3, result.Series[0].Values.Length);
        Assert.Equal([10.0, 20.0, 30.0], result.Series[0].Values);
    }

    [Fact]
    public void UpdateChartData_ScatterChart_UpdatesYValues()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Scatter",
                    Series = [new TestSeriesDefinition { Name = "Dataset", Values = [5.0, 10.0, 15.0] }]
                }
            ]
        });

        Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: null, Values: [50.0, 100.0, 75.0])
        ]);

        var readBack = Service.GetChartData(path, slideNumber: 1);
        Assert.Equal("Scatter", readBack.ChartType);
        Assert.Equal([50.0, 100.0, 75.0], readBack.Series[0].Values);
    }

    // ── Selector validation ───────────────────────────────────────────────────────

    [Fact]
    public void GetChartData_MultipleChartsNoSelector_Throws()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition { Name = "Chart A", ChartType = "Column", Series = [new TestSeriesDefinition { Values = [1.0] }] },
                new TestChartDefinition { Name = "Chart B", ChartType = "Bar", Series = [new TestSeriesDefinition { Values = [2.0] }] }
            ]
        });

        var ex = Assert.Throws<InvalidOperationException>(() =>
            Service.GetChartData(path, slideNumber: 1));
        Assert.Contains("Specify chartName or chartIndex", ex.Message);
    }

    [Fact]
    public void UpdateChartData_MultipleChartsNoSelector_Throws()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition { Name = "Chart A", ChartType = "Column", Series = [new TestSeriesDefinition { Values = [1.0] }] },
                new TestChartDefinition { Name = "Chart B", ChartType = "Bar", Series = [new TestSeriesDefinition { Values = [2.0] }] }
            ]
        });

        Assert.Throws<InvalidOperationException>(() =>
            Service.UpdateChartData(path, slideNumber: 1, [
                new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: null, Values: [9.0])
            ]));
    }

    // ── Length mismatch validation ────────────────────────────────────────────────

    [Fact]
    public void UpdateChartData_MismatchedCategoriesAndValues_ReturnsFailure()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Column",
                    Categories = ["A", "B"],
                    Series = [new TestSeriesDefinition { Name = "S", Values = [1.0, 2.0] }]
                }
            ]
        });

        var result = Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: ["X", "Y", "Z"], Values: [1.0, 2.0])
        ]);

        Assert.False(result.Success);
        Assert.Contains("same length", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UpdateChartData_CategoriesOnlyNoMismatch_Succeeds()
    {
        // Providing only Categories (no Values) should not trigger length validation
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    ChartType = "Column",
                    Categories = ["A", "B"],
                    Series = [new TestSeriesDefinition { Name = "S", Values = [1.0, 2.0] }]
                }
            ]
        });

        var result = Service.UpdateChartData(path, slideNumber: 1, [
            new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: ["X", "Y", "Z"], Values: null)
        ]);

        Assert.True(result.Success);
    }
}
