using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxMcp.Models;

namespace PptxMcp.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Read or update data in an existing chart shape on a slide without changing its styling or formatting.
    /// Available actions:
    /// - Read: Return chart type, series names, categories, and data values from an existing chart.
    /// - Update: Replace series data values (and optionally names and categories) while preserving all chart formatting.
    /// Supported chart types: Column, Bar, Line, Pie, Area, Scatter, and their 3D and Doughnut variants.
    /// Locate the chart by chartName (case-insensitive) or chartIndex (zero-based). When the slide has only one chart, both may be omitted.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="action">The chart data operation to perform: Read or Update.</param>
    /// <param name="slideNumber">1-based slide number containing the chart.</param>
    /// <param name="chartName">Optional chart shape name to match (case-insensitive). Takes precedence over chartIndex.</param>
    /// <param name="chartIndex">Optional zero-based index among chart shapes on the slide. Used when chartName is not provided.</param>
    /// <param name="updates">Array of series updates for the Update action. Each entry targets a series by zero-based SeriesIndex and may supply a new SeriesName, Categories array, and/or Values array.</param>
    [McpServerTool(Title = "Chart Data")]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["Read","Update"]""")]
    public partial Task<string> pptx_chart_data(
        string filePath,
        ChartDataAction action,
        int slideNumber,
        string? chartName = null,
        int? chartIndex = null,
        ChartSeriesUpdate[]? updates = null)
    {
        return action switch
        {
            ChartDataAction.Read => ExecuteToolStructured(filePath,
                () => _service.GetChartData(filePath, slideNumber, chartName, chartIndex),
                error => new ChartDataResult(
                    Success: false,
                    SlideNumber: slideNumber,
                    ChartName: null,
                    MatchedBy: null,
                    ChartType: null,
                    SeriesCount: 0,
                    Series: [],
                    Message: error)),

            ChartDataAction.Update => ExecuteToolStructured(filePath,
                () =>
                {
                    if (updates is null || updates.Length == 0)
                        throw new ArgumentException("updates is required for the Update action and must contain at least one entry.");
                    return _service.UpdateChartData(filePath, slideNumber, updates, chartName, chartIndex);
                },
                error => new ChartUpdateResult(
                    Success: false,
                    SlideNumber: slideNumber,
                    ChartName: null,
                    MatchedBy: null,
                    ChartType: null,
                    SeriesUpdated: 0,
                    Message: error)),

            _ => Task.FromResult(JsonSerializer.Serialize(
                new { Success = false, Message = $"Unknown action: {action}. Valid actions: Read, Update." },
                IndentedJson))
        };
    }
}
