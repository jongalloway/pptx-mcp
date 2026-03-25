namespace PptxTools.Models;

/// <summary>Describes updates to apply to a single chart series for the Update action of pptx_chart_data.</summary>
/// <param name="SeriesIndex">Zero-based index of the series to update.</param>
/// <param name="SeriesName">Optional new display name for the series. Omit to keep the existing name.</param>
/// <param name="Categories">Optional new category labels. Omit to keep existing categories. Must match the length of Values when both are provided.</param>
/// <param name="Values">Optional new numeric data values. Omit to keep existing values.</param>
public record ChartSeriesUpdate(
    int SeriesIndex,
    string? SeriesName,
    string[]? Categories,
    double[]? Values);
