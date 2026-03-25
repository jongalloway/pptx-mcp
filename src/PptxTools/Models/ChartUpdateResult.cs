namespace PptxTools.Models;

/// <summary>Structured result for the Update action of pptx_chart_data.</summary>
/// <param name="Success">True when chart data was updated successfully.</param>
/// <param name="SlideNumber">1-based slide number that was targeted.</param>
/// <param name="ChartName">Name of the matched chart shape.</param>
/// <param name="MatchedBy">How the chart was located: chartName, chartIndex, or onlyChart.</param>
/// <param name="ChartType">Detected chart type (e.g. Column, Bar, Line, Pie, Area, Scatter).</param>
/// <param name="SeriesUpdated">Number of series that were updated.</param>
/// <param name="Message">Human-readable status message describing the outcome.</param>
public record ChartUpdateResult(
    bool Success,
    int SlideNumber,
    string? ChartName,
    string? MatchedBy,
    string? ChartType,
    int SeriesUpdated,
    string Message);
