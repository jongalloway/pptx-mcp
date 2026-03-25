namespace PptxTools.Models;

/// <summary>Structured result for the Read action of pptx_chart_data.</summary>
/// <param name="Success">True when chart data was read successfully.</param>
/// <param name="SlideNumber">1-based slide number that was targeted.</param>
/// <param name="ChartName">Name of the matched chart shape.</param>
/// <param name="MatchedBy">How the chart was located: chartName, chartIndex, or onlyChart.</param>
/// <param name="ChartType">Detected chart type (e.g. Column, Bar, Line, Pie, Area, Scatter).</param>
/// <param name="SeriesCount">Number of data series in the chart.</param>
/// <param name="Series">Array of series data including names, categories, and values.</param>
/// <param name="Message">Human-readable status message describing the outcome.</param>
public record ChartDataResult(
    bool Success,
    int SlideNumber,
    string? ChartName,
    string? MatchedBy,
    string? ChartType,
    int SeriesCount,
    ChartSeriesData[] Series,
    string Message);
