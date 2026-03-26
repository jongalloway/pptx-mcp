using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxTools.Models;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace PptxTools.Services;

public partial class PresentationService
{
    /// <summary>Export a presentation to a structured JSON-ready object.</summary>
    public PresentationExport ExportJson(string filePath, ExportJsonAction action)
    {
        if (action == ExportJsonAction.SchemaOnly)
            return BuildSchemaDescription();

        using var doc = PresentationDocument.Open(filePath, false);
        var presentationPart = doc.PresentationPart;
        if (presentationPart is null)
            return MakeExportError(action, filePath, "Presentation part not found.");

        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<SlideId>().ToList() ?? [];
        int slideCount = slideIds.Count;

        var metadata = action is ExportJsonAction.SlidesOnly
            ? null
            : BuildMetadataExport(doc);

        var slides = action is ExportJsonAction.MetadataOnly
            ? null
            : BuildSlideExports(doc, presentationPart, slideIds);

        var slidesSummary = slides is not null ? $"{slides.Count} slide(s)" : "excluded";
        var metaSummary = metadata is not null ? "metadata included" : "metadata excluded";

        return new PresentationExport(
            Success: true,
            Action: action.ToString(),
            FilePath: filePath,
            Metadata: metadata,
            SlideCount: slideCount,
            Slides: slides,
            Schema: null,
            Message: $"Exported {slidesSummary}, {metaSummary}.");
    }

    private static PresentationMetadataExport BuildMetadataExport(PresentationDocument doc)
    {
        var props = doc.PackageProperties;
        return new PresentationMetadataExport(
            Title: props.Title,
            Creator: props.Creator,
            Created: props.Created?.ToString("o"),
            Modified: props.Modified?.ToString("o"),
            Subject: props.Subject,
            Keywords: props.Keywords,
            Description: props.Description,
            LastModifiedBy: props.LastModifiedBy,
            Category: props.Category);
    }

    private List<SlideExport> BuildSlideExports(
        PresentationDocument doc,
        PresentationPart presentationPart,
        List<SlideId> slideIds)
    {
        var slides = new List<SlideExport>(slideIds.Count);
        for (int i = 0; i < slideIds.Count; i++)
        {
            var slidePart = GetSlidePart(doc, slideIds, i);
            slides.Add(BuildSlideExport(presentationPart, slidePart, i));
        }
        return slides;
    }

    private SlideExport BuildSlideExport(PresentationPart presentationPart, SlidePart slidePart, int slideIndex)
    {
        var content = GetSlideContent(presentationPart, slidePart, slideIndex);
        var title = ExtractSlideTitle(content);
        var notes = GetSlideNotes(slidePart);

        // Pre-build chart lookup keyed by shape name
        var chartLookup = BuildChartLookup(slidePart);

        var shapes = new List<ShapeExport>();
        foreach (var shape in content.Shapes)
        {
            TableExportData? tableData = null;
            ImageExport? imageData = null;
            ChartExport? chartData = null;

            if (shape.ShapeType == "Table" && shape.TableRows is not null)
            {
                var rows = shape.TableRows;
                int colCount = rows.Count > 0 ? rows[0].Count : 0;
                tableData = new TableExportData(
                    RowCount: rows.Count,
                    ColumnCount: colCount,
                    Cells: rows);
            }

            if (shape.ShapeType == "Picture")
            {
                imageData = ExtractImageExport(slidePart, shape);
            }

            if (chartLookup.TryGetValue(shape.Name, out var chart))
            {
                chartData = chart;
            }

            shapes.Add(new ShapeExport(
                ShapeId: shape.ShapeId,
                Name: shape.Name,
                ShapeType: shape.ShapeType,
                X: shape.X,
                Y: shape.Y,
                Width: shape.Width,
                Height: shape.Height,
                IsPlaceholder: shape.IsPlaceholder,
                PlaceholderType: shape.PlaceholderType,
                Text: shape.Text,
                Paragraphs: shape.Paragraphs,
                Table: tableData,
                Image: imageData,
                Chart: chartData));
        }

        return new SlideExport(
            SlideNumber: slideIndex + 1,
            SlideIndex: slideIndex,
            Title: title,
            SlideWidthEmu: content.SlideWidthEmu,
            SlideHeightEmu: content.SlideHeightEmu,
            Shapes: shapes,
            SpeakerNotes: notes);
    }

    private static string? ExtractSlideTitle(SlideContent content)
    {
        foreach (var shape in content.Shapes)
        {
            if (shape.IsPlaceholder &&
                shape.PlaceholderType is "title" or "ctrTitle" &&
                !string.IsNullOrWhiteSpace(shape.Text))
            {
                return shape.Text;
            }
        }
        return null;
    }

    private static ImageExport? ExtractImageExport(SlidePart slidePart, ShapeContent shape)
    {
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree is null) return null;

        foreach (var pic in shapeTree.Elements<Picture>())
        {
            var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
            if (name != shape.Name) continue;

            var embed = pic.BlipFill?.Blip?.Embed?.Value;
            if (embed is null) continue;

            var imagePart = slidePart.GetPartById(embed);
            var contentType = imagePart?.ContentType ?? "unknown";
            var format = contentType.Split('/').LastOrDefault()?.ToUpperInvariant() ?? "UNKNOWN";

            return new ImageExport(
                ShapeName: shape.Name,
                ContentType: contentType,
                ImageFormat: format,
                RelationshipId: embed,
                WidthEmu: shape.Width,
                HeightEmu: shape.Height);
        }

        return null;
    }

    private Dictionary<string, ChartExport> BuildChartLookup(SlidePart slidePart)
    {
        var lookup = new Dictionary<string, ChartExport>();
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree is null) return lookup;

        foreach (var graphicFrame in shapeTree.Elements<GraphicFrame>())
        {
            var graphicData = graphicFrame.Graphic?.GraphicData;
            if (graphicData?.Uri?.Value != ChartGraphicDataUri) continue;

            var chartRef = graphicData.GetFirstChild<C.ChartReference>();
            if (chartRef?.Id?.Value is null) continue;

            try
            {
                var chartPart = (ChartPart)slidePart.GetPartById(chartRef.Id.Value);
                var shapeName = graphicFrame.NonVisualGraphicFrameProperties?
                    .NonVisualDrawingProperties?.Name?.Value ?? "Chart";

                var (chartType, seriesData) = ExtractChartData(chartPart);

                var seriesExports = seriesData.Select(s => new ChartSeriesExport(
                    SeriesIndex: s.SeriesIndex,
                    SeriesName: s.SeriesName,
                    Categories: s.Categories,
                    Values: s.Values)).ToList();

                lookup[shapeName] = new ChartExport(
                    ShapeName: shapeName,
                    ChartType: chartType,
                    SeriesCount: seriesExports.Count,
                    Series: seriesExports);
            }
            catch
            {
                // Skip charts that can't be read
            }
        }

        return lookup;
    }

    private static PresentationExport BuildSchemaDescription()
    {
        const string schema = "PresentationExport { Success, Action, FilePath, " +
            "Metadata { Title, Creator, Created, Modified, Subject, Keywords, Description, LastModifiedBy, Category }, " +
            "SlideCount, Slides[] { SlideNumber, SlideIndex, Title, SlideWidthEmu, SlideHeightEmu, " +
            "Shapes[] { ShapeId, Name, ShapeType, X, Y, Width, Height, IsPlaceholder, PlaceholderType, Text, Paragraphs, " +
            "Table { RowCount, ColumnCount, Cells[][] }, " +
            "Image { ContentType, ImageFormat, RelationshipId, WidthEmu, HeightEmu }, " +
            "Chart { ChartType, SeriesCount, Series[] { SeriesIndex, SeriesName, Categories[], Values[] } } }, " +
            "SpeakerNotes }, Schema, Message }";

        return new PresentationExport(
            Success: true,
            Action: "SchemaOnly",
            FilePath: null,
            Metadata: null,
            SlideCount: 0,
            Slides: null,
            Schema: schema,
            Message: "Schema description returned. No file was read.");
    }

    private static PresentationExport MakeExportError(ExportJsonAction action, string? filePath, string message) =>
        new(
            Success: false,
            Action: action.ToString(),
            FilePath: filePath,
            Metadata: null,
            SlideCount: 0,
            Slides: null,
            Schema: null,
            Message: message);
}
