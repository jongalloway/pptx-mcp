using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxMcp.Models;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Services;

public partial class PresentationService
{
    private const string ChartGraphicDataUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public ChartDataResult GetChartData(string filePath, int slideNumber, string? chartName = null, int? chartIndex = null)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var slideIds = GetSlideIds(doc);
        var slidePart = GetSlidePart(doc, slideIds, slideNumber - 1);
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        var chartFrames = GetChartFrames(shapeTree);

        if (chartFrames.Count == 0)
            return new ChartDataResult(false, slideNumber, null, null, null, 0, [],
                $"Slide {slideNumber} has no charts.");

        var (targetFrame, matchedBy) = ResolveChartFrame(chartFrames, chartName, chartIndex, slideNumber);
        var (chartPart, resolvedName) = GetChartPart(slidePart, targetFrame);

        var (chartType, series) = ExtractChartData(chartPart);

        return new ChartDataResult(
            Success: true,
            SlideNumber: slideNumber,
            ChartName: resolvedName,
            MatchedBy: matchedBy,
            ChartType: chartType,
            SeriesCount: series.Length,
            Series: series,
            Message: $"Read {series.Length} series from '{resolvedName}' (type: {chartType}) on slide {slideNumber}.");
    }

    public ChartUpdateResult UpdateChartData(
        string filePath,
        int slideNumber,
        ChartSeriesUpdate[] updates,
        string? chartName = null,
        int? chartIndex = null)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var slideIds = GetSlideIds(doc);
        var slidePart = GetSlidePart(doc, slideIds, slideNumber - 1);
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        var chartFrames = GetChartFrames(shapeTree);

        if (chartFrames.Count == 0)
            return new ChartUpdateResult(false, slideNumber, null, null, null, 0,
                $"Slide {slideNumber} has no charts.");

        var (targetFrame, matchedBy) = ResolveChartFrame(chartFrames, chartName, chartIndex, slideNumber);
        var (chartPart, resolvedName) = GetChartPart(slidePart, targetFrame);

        var (chartType, seriesElements) = GetChartSeriesElements(chartPart);
        int seriesUpdated = 0;

        foreach (var update in updates)
        {
            if (update.SeriesIndex < 0 || update.SeriesIndex >= seriesElements.Count)
                continue;

            var seriesElement = seriesElements[update.SeriesIndex];
            ApplySeriesUpdate(seriesElement, update);
            seriesUpdated++;
        }

        chartPart.ChartSpace.Save();

        return new ChartUpdateResult(
            Success: true,
            SlideNumber: slideNumber,
            ChartName: resolvedName,
            MatchedBy: matchedBy,
            ChartType: chartType,
            SeriesUpdated: seriesUpdated,
            Message: $"Updated {seriesUpdated} series in '{resolvedName}' (type: {chartType}) on slide {slideNumber}.");
    }

    // ── Private helpers ──────────────────────────────────────────────────────────

    private static List<P.GraphicFrame> GetChartFrames(ShapeTree shapeTree) =>
        shapeTree.Elements<P.GraphicFrame>()
            .Where(gf => gf.Graphic?.GraphicData?.Uri?.Value == ChartGraphicDataUri)
            .ToList();

    private static (P.GraphicFrame Frame, string MatchedBy) ResolveChartFrame(
        List<P.GraphicFrame> chartFrames,
        string? chartName,
        int? chartIndex,
        int slideNumber)
    {
        if (chartFrames.Count == 1 && chartName is null && chartIndex is null)
            return (chartFrames[0], "onlyChart");

        if (!string.IsNullOrWhiteSpace(chartName))
        {
            var frame = chartFrames.FirstOrDefault(gf =>
                string.Equals(
                    gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value,
                    chartName,
                    StringComparison.OrdinalIgnoreCase));

            if (frame is null)
            {
                var available = string.Join(", ",
                    chartFrames.Select((gf, i) =>
                        $"{i}:{gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "(unnamed)"}"));
                throw new InvalidOperationException(
                    $"No chart named '{chartName}' found on slide {slideNumber}. Available charts: {available}");
            }

            return (frame, "chartName");
        }

        if (chartIndex.HasValue)
        {
            if (chartIndex.Value < 0 || chartIndex.Value >= chartFrames.Count)
                throw new ArgumentOutOfRangeException(nameof(chartIndex),
                    $"Chart index {chartIndex.Value} is out of range. Slide {slideNumber} has {chartFrames.Count} chart(s).");

            return (chartFrames[chartIndex.Value], "chartIndex");
        }

        // Multiple charts, no selector — default to first
        return (chartFrames[0], "chartIndex");
    }

    private static (ChartPart ChartPart, string? ResolvedName) GetChartPart(SlidePart slidePart, P.GraphicFrame frame)
    {
        var chartRef = frame.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>();
        var relId = chartRef?.Id?.Value;

        if (relId is null)
            throw new InvalidOperationException("Chart shape does not contain a valid chart reference.");

        if (slidePart.GetPartById(relId) is not ChartPart chartPart)
            throw new InvalidOperationException($"Could not load ChartPart for relationship '{relId}'.");

        var resolvedName = frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value;
        return (chartPart, resolvedName);
    }

    private static (string ChartType, ChartSeriesData[] Series) ExtractChartData(ChartPart chartPart)
    {
        var (chartType, seriesElements) = GetChartSeriesElements(chartPart);
        var series = seriesElements
            .Select((el, i) => ExtractSeriesData(el, i))
            .ToArray();
        return (chartType, series);
    }

    private static (string ChartType, List<OpenXmlCompositeElement> SeriesElements) GetChartSeriesElements(ChartPart chartPart)
    {
        var plotArea = chartPart.ChartSpace
            .GetFirstChild<C.Chart>()?
            .GetFirstChild<C.PlotArea>();

        if (plotArea is null)
            return ("Unknown", []);

        // Check supported chart types in priority order
        if (plotArea.GetFirstChild<C.BarChart>() is C.BarChart barChart)
        {
            var dir = barChart.GetFirstChild<C.BarDirection>()?.Val?.Value;
            var chartType = dir == C.BarDirectionValues.Column ? "Column" : "Bar";
            return (chartType, barChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>().ToList());
        }

        if (plotArea.GetFirstChild<C.Bar3DChart>() is C.Bar3DChart bar3DChart)
        {
            var dir = bar3DChart.GetFirstChild<C.BarDirection>()?.Val?.Value;
            var chartType = dir == C.BarDirectionValues.Column ? "Column3D" : "Bar3D";
            return (chartType, bar3DChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>().ToList());
        }

        if (plotArea.GetFirstChild<C.LineChart>() is C.LineChart lineChart)
            return ("Line", lineChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>().ToList());

        if (plotArea.GetFirstChild<C.Line3DChart>() is C.Line3DChart line3DChart)
            return ("Line3D", line3DChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>().ToList());

        if (plotArea.GetFirstChild<C.PieChart>() is C.PieChart pieChart)
            return ("Pie", pieChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>().ToList());

        if (plotArea.GetFirstChild<C.Pie3DChart>() is C.Pie3DChart pie3DChart)
            return ("Pie3D", pie3DChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>().ToList());

        if (plotArea.GetFirstChild<C.DoughnutChart>() is C.DoughnutChart doughnutChart)
            return ("Doughnut", doughnutChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>().ToList());

        if (plotArea.GetFirstChild<C.AreaChart>() is C.AreaChart areaChart)
            return ("Area", areaChart.Elements<C.AreaChartSeries>().Cast<OpenXmlCompositeElement>().ToList());

        if (plotArea.GetFirstChild<C.Area3DChart>() is C.Area3DChart area3DChart)
            return ("Area3D", area3DChart.Elements<C.AreaChartSeries>().Cast<OpenXmlCompositeElement>().ToList());

        if (plotArea.GetFirstChild<C.ScatterChart>() is C.ScatterChart scatterChart)
            return ("Scatter", scatterChart.Elements<C.ScatterChartSeries>().Cast<OpenXmlCompositeElement>().ToList());

        return ("Unknown", []);
    }

    private static ChartSeriesData ExtractSeriesData(OpenXmlCompositeElement seriesElement, int index)
    {
        var seriesName = GetSeriesNameFromElement(seriesElement);
        var categories = GetStringCacheValues(seriesElement.GetFirstChild<C.CategoryAxisData>());
        var values = GetNumericCacheValues(seriesElement.GetFirstChild<C.Values>());
        return new ChartSeriesData(index, seriesName, categories, values);
    }

    private static string? GetSeriesNameFromElement(OpenXmlCompositeElement seriesElement)
    {
        var seriesText = seriesElement.GetFirstChild<C.SeriesText>();
        if (seriesText is null) return null;

        // Try StringReference cache first
        var strRef = seriesText.GetFirstChild<C.StringReference>();
        if (strRef is not null)
        {
            var strCache = strRef.GetFirstChild<C.StringCache>();
            var pt = strCache?.Elements<C.StringPoint>().FirstOrDefault();
            if (pt?.NumericValue?.InnerText is string cached && cached.Length > 0)
                return cached;
        }

        // Try inline NumericValue
        return seriesText.GetFirstChild<C.NumericValue>()?.InnerText;
    }

    private static string[] GetStringCacheValues(C.CategoryAxisData? catAxisData)
    {
        if (catAxisData is null) return [];

        // String reference (most common for category labels)
        var strCache = catAxisData.GetFirstChild<C.StringReference>()?.GetFirstChild<C.StringCache>();
        if (strCache is not null)
            return strCache.Elements<C.StringPoint>()
                .OrderBy(pt => pt.Index?.Value ?? 0)
                .Select(pt => pt.NumericValue?.InnerText ?? string.Empty)
                .ToArray();

        // Numeric reference (e.g. year numbers as categories)
        var numCache = catAxisData.GetFirstChild<C.NumberReference>()?.GetFirstChild<C.NumberingCache>();
        if (numCache is not null)
            return numCache.Elements<C.NumericPoint>()
                .OrderBy(pt => pt.Index?.Value ?? 0)
                .Select(pt => pt.NumericValue?.InnerText ?? string.Empty)
                .ToArray();

        return [];
    }

    private static double[] GetNumericCacheValues(C.Values? valuesElement)
    {
        if (valuesElement is null) return [];

        var numCache = valuesElement.GetFirstChild<C.NumberReference>()?.GetFirstChild<C.NumberingCache>();
        if (numCache is null) return [];

        return numCache.Elements<C.NumericPoint>()
            .OrderBy(pt => pt.Index?.Value ?? 0)
            .Select(pt => double.TryParse(pt.NumericValue?.InnerText, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : 0.0)
            .ToArray();
    }

    private static void ApplySeriesUpdate(OpenXmlCompositeElement seriesElement, ChartSeriesUpdate update)
    {
        if (update.SeriesName is not null)
            UpdateSeriesName(seriesElement, update.SeriesName);

        if (update.Categories is not null)
            UpdateCategoryCache(seriesElement, update.Categories);

        if (update.Values is not null)
            UpdateValueCache(seriesElement, update.Values);
    }

    private static void UpdateSeriesName(OpenXmlCompositeElement seriesElement, string newName)
    {
        var seriesText = seriesElement.GetFirstChild<C.SeriesText>();
        if (seriesText is null) return;

        var strRef = seriesText.GetFirstChild<C.StringReference>();
        if (strRef is not null)
        {
            var strCache = strRef.GetFirstChild<C.StringCache>();
            if (strCache is null)
            {
                strCache = new C.StringCache();
                strRef.Append(strCache);
            }

            // Remove existing points and rebuild
            foreach (var pt in strCache.Elements<C.StringPoint>().ToList())
                pt.Remove();
            strCache.GetFirstChild<C.PointCount>()?.Remove();

            strCache.Append(new C.PointCount { Val = 1U });
            strCache.Append(new C.StringPoint { Index = 0U, NumericValue = new C.NumericValue(newName) });
            return;
        }

        // Inline NumericValue fallback
        var inlineVal = seriesText.GetFirstChild<C.NumericValue>();
        if (inlineVal is not null)
            inlineVal.Text = newName;
        else
            seriesText.Append(new C.NumericValue(newName));
    }

    private static void UpdateCategoryCache(OpenXmlCompositeElement seriesElement, string[] categories)
    {
        var catAxisData = seriesElement.GetFirstChild<C.CategoryAxisData>();
        if (catAxisData is null) return;

        var strRef = catAxisData.GetFirstChild<C.StringReference>();
        if (strRef is not null)
        {
            var strCache = strRef.GetFirstChild<C.StringCache>();
            if (strCache is null)
            {
                strCache = new C.StringCache();
                strRef.Append(strCache);
            }

            RebuildStringCache(strCache, categories);
            return;
        }

        var numRef = catAxisData.GetFirstChild<C.NumberReference>();
        if (numRef is not null)
        {
            // Convert numeric ref to string ref when updating with string categories
            numRef.Remove();
            var newStrRef = new C.StringReference(
                new C.Formula(string.Empty),
                BuildStringCache(categories));
            catAxisData.Append(newStrRef);
        }
    }

    private static void UpdateValueCache(OpenXmlCompositeElement seriesElement, double[] values)
    {
        var valuesElement = seriesElement.GetFirstChild<C.Values>();
        if (valuesElement is null) return;

        var numRef = valuesElement.GetFirstChild<C.NumberReference>();
        if (numRef is null) return;

        var numCache = numRef.GetFirstChild<C.NumberingCache>();
        if (numCache is null)
        {
            numCache = new C.NumberingCache();
            numRef.Append(numCache);
        }

        RebuildNumberingCache(numCache, values);
    }

    private static void RebuildStringCache(C.StringCache strCache, string[] values)
    {
        foreach (var pt in strCache.Elements<C.StringPoint>().ToList())
            pt.Remove();
        strCache.GetFirstChild<C.PointCount>()?.Remove();

        strCache.Append(new C.PointCount { Val = (uint)values.Length });
        for (int i = 0; i < values.Length; i++)
            strCache.Append(new C.StringPoint { Index = (uint)i, NumericValue = new C.NumericValue(values[i]) });
    }

    private static C.StringCache BuildStringCache(string[] values)
    {
        var cache = new C.StringCache(new C.PointCount { Val = (uint)values.Length });
        for (int i = 0; i < values.Length; i++)
            cache.Append(new C.StringPoint { Index = (uint)i, NumericValue = new C.NumericValue(values[i]) });
        return cache;
    }

    private static void RebuildNumberingCache(C.NumberingCache numCache, double[] values)
    {
        foreach (var pt in numCache.Elements<C.NumericPoint>().ToList())
            pt.Remove();
        numCache.GetFirstChild<C.PointCount>()?.Remove();

        numCache.Append(new C.PointCount { Val = (uint)values.Length });
        for (int i = 0; i < values.Length; i++)
            numCache.Append(new C.NumericPoint
            {
                Index = (uint)i,
                NumericValue = new C.NumericValue(values[i].ToString(System.Globalization.CultureInfo.InvariantCulture))
            });
    }
}
