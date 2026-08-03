namespace PptxTools.Models;

/// <summary>Data for a single chart series returned by the Read action of pptx_chart_data.</summary>
/// <param name="SeriesIndex">Zero-based index of the series within the chart.</param>
/// <param name="SeriesName">Display name of the series, or null if unnamed.</param>
/// <param name="Categories">Category labels for the horizontal axis. Empty for chart types without categories.</param>
/// <param name="Values">Numeric data values for the series.</param>
public record ChartSeriesData(
    int SeriesIndex,
    string? SeriesName,
    string[] Categories,
    double[] Values);
