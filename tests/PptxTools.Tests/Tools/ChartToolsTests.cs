using System.Text.Json;

namespace PptxTools.Tests.Tools;

[Trait("Category", "Integration")]
public class ChartToolsTests : PptxTestBase
{
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public ChartToolsTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    // ── File-not-found: structured error ────────────────────────────────────────

    [Fact]
    public async Task pptx_chart_data_Read_FileNotFound_ReturnsStructuredError()
    {
        var fakePath = Path.Join(Path.GetTempPath(), "nonexistent-chart-test.pptx");

        var result = await _tools.pptx_chart_data(fakePath, ChartDataAction.Read, slideNumber: 1);

        var parsed = JsonSerializer.Deserialize<ChartDataResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("File not found", parsed.Message);
    }

    [Fact]
    public async Task pptx_chart_data_Update_FileNotFound_ReturnsStructuredError()
    {
        var fakePath = Path.Join(Path.GetTempPath(), "nonexistent-chart-test.pptx");

        var result = await _tools.pptx_chart_data(fakePath, ChartDataAction.Update, slideNumber: 1,
            updates: [new ChartSeriesUpdate(0, null, null, [1.0])]);

        var parsed = JsonSerializer.Deserialize<ChartUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("File not found", parsed.Message);
    }

    // ── Read action ──────────────────────────────────────────────────────────────

    [Fact]
    public async Task pptx_chart_data_Read_ColumnChart_ReturnsJson()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    Name = "Revenue Chart",
                    ChartType = "Column",
                    Categories = ["Q1", "Q2", "Q3"],
                    Series = [new TestSeriesDefinition { Name = "Revenue", Values = [100.0, 200.0, 150.0] }]
                }
            ]
        });

        var result = await _tools.pptx_chart_data(path, ChartDataAction.Read, slideNumber: 1);

        var parsed = JsonSerializer.Deserialize<ChartDataResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("Revenue Chart", parsed.ChartName);
        Assert.Equal("Column", parsed.ChartType);
        Assert.Equal(1, parsed.SeriesCount);
        Assert.Single(parsed.Series);
        Assert.Equal("Revenue", parsed.Series[0].SeriesName);
        Assert.Equal(["Q1", "Q2", "Q3"], parsed.Series[0].Categories);
        Assert.Equal([100.0, 200.0, 150.0], parsed.Series[0].Values);
    }

    [Fact]
    public async Task pptx_chart_data_Read_NoCharts_ReturnsFailure()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_chart_data(path, ChartDataAction.Read, slideNumber: 1);

        var parsed = JsonSerializer.Deserialize<ChartDataResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    // ── Update action ────────────────────────────────────────────────────────────

    [Fact]
    public async Task pptx_chart_data_Update_NoUpdates_ReturnsError()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts = [new TestChartDefinition { ChartType = "Column", Series = [new TestSeriesDefinition { Values = [1.0] }] }]
        });

        var result = await _tools.pptx_chart_data(path, ChartDataAction.Update, slideNumber: 1, updates: null);

        var parsed = JsonSerializer.Deserialize<ChartUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("updates is required", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task pptx_chart_data_Update_UpdatesValues_ReturnsSuccess()
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

        var result = await _tools.pptx_chart_data(path, ChartDataAction.Update, slideNumber: 1,
            updates: [new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: null, Categories: null, Values: [10.0, 20.0, 30.0])]);

        var parsed = JsonSerializer.Deserialize<ChartUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal(1, parsed.SeriesUpdated);
        Assert.Equal("Column", parsed.ChartType);

        // Verify the update is reflected when reading back
        var readResult = await _tools.pptx_chart_data(path, ChartDataAction.Read, slideNumber: 1);
        var readParsed = JsonSerializer.Deserialize<ChartDataResult>(readResult);
        Assert.NotNull(readParsed);
        Assert.Equal([10.0, 20.0, 30.0], readParsed.Series[0].Values);
    }

    [Fact]
    public async Task pptx_chart_data_Update_ByChartName_ReturnsSuccess()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            Charts =
            [
                new TestChartDefinition
                {
                    Name = "Target Chart",
                    ChartType = "Column",
                    Series = [new TestSeriesDefinition { Name = "S1", Values = [1.0, 2.0] }]
                }
            ]
        });

        var result = await _tools.pptx_chart_data(path, ChartDataAction.Update, slideNumber: 1,
            chartName: "Target Chart",
            updates: [new ChartSeriesUpdate(SeriesIndex: 0, SeriesName: "Updated Series", Categories: null, Values: null)]);

        var parsed = JsonSerializer.Deserialize<ChartUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("chartName", parsed.MatchedBy);
        Assert.Equal("Target Chart", parsed.ChartName);
    }

    // ── Unknown action ───────────────────────────────────────────────────────────

    [Fact]
    public async Task pptx_chart_data_UnknownAction_ReturnsErrorJson()
    {
        var path = CreateMinimalPptx();

        // Cast an invalid int to the enum to simulate an unknown action value
        var result = await _tools.pptx_chart_data(path, (ChartDataAction)99, slideNumber: 1);

        Assert.Contains("Unknown action", result);
    }
}
