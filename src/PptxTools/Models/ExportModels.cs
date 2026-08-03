namespace PptxTools.Models;

/// <summary>Actions for the consolidated pptx_export_json tool.</summary>
public enum ExportJsonAction
{
    /// <summary>Export the full presentation including metadata, all slides, and all content.</summary>
    Full,

    /// <summary>Export only the slide content (shapes, tables, charts, images, notes).</summary>
    SlidesOnly,

    /// <summary>Export only presentation-level metadata.</summary>
    MetadataOnly,

    /// <summary>Return the JSON schema description without reading any file.</summary>
    SchemaOnly
}

/// <summary>Top-level export result wrapping the full presentation structure.</summary>
/// <param name="Success">Whether the export completed without errors.</param>
/// <param name="Action">The export action that was performed.</param>
/// <param name="FilePath">Path to the source presentation.</param>
/// <param name="Metadata">Presentation-level metadata. Null when action is SlidesOnly.</param>
/// <param name="SlideCount">Total number of slides in the presentation.</param>
/// <param name="Slides">Exported slide structures. Null when action is MetadataOnly.</param>
/// <param name="Schema">Schema description string. Populated only for SchemaOnly action.</param>
/// <param name="Message">Human-readable summary or error message.</param>
public record PresentationExport(
    bool Success,
    string Action,
    string? FilePath,
    PresentationMetadataExport? Metadata,
    int SlideCount,
    IReadOnlyList<SlideExport>? Slides,
    string? Schema,
    string Message);

/// <summary>Presentation metadata for export.</summary>
/// <param name="Title">Document title.</param>
/// <param name="Creator">Author / creator.</param>
/// <param name="Created">Creation date in ISO 8601 format.</param>
/// <param name="Modified">Last modified date in ISO 8601 format.</param>
/// <param name="Subject">Subject field.</param>
/// <param name="Keywords">Keywords / tags.</param>
/// <param name="Description">Description or comments.</param>
/// <param name="LastModifiedBy">Last person who saved changes.</param>
/// <param name="Category">Category field.</param>
public record PresentationMetadataExport(
    string? Title,
    string? Creator,
    string? Created,
    string? Modified,
    string? Subject,
    string? Keywords,
    string? Description,
    string? LastModifiedBy,
    string? Category);

/// <summary>Exported slide with all content types.</summary>
/// <param name="SlideNumber">1-based slide number.</param>
/// <param name="SlideIndex">Zero-based slide index.</param>
/// <param name="Title">Slide title extracted from the title placeholder. Null if absent.</param>
/// <param name="SlideWidthEmu">Slide width in EMUs.</param>
/// <param name="SlideHeightEmu">Slide height in EMUs.</param>
/// <param name="Shapes">All shapes on the slide with embedded table/image/chart data.</param>
/// <param name="SpeakerNotes">Speaker notes text. Null if no notes exist.</param>
public record SlideExport(
    int SlideNumber,
    int SlideIndex,
    string? Title,
    long SlideWidthEmu,
    long SlideHeightEmu,
    IReadOnlyList<ShapeExport> Shapes,
    string? SpeakerNotes)
{
    /// <summary>Convenience: chart data aggregated from chart shapes.</summary>
    public IReadOnlyList<ChartExport> Charts =>
        Shapes.Where(s => s.Chart is not null).Select(s => s.Chart!).ToList();

    /// <summary>Convenience: image metadata aggregated from picture shapes.</summary>
    public IReadOnlyList<ImageExport> Images =>
        Shapes.Where(s => s.Image is not null).Select(s => s.Image!).ToList();
}

/// <summary>Exported shape data with optional embedded sub-type content.</summary>
/// <param name="ShapeId">Numeric shape ID from OpenXML.</param>
/// <param name="Name">Shape name.</param>
/// <param name="ShapeType">Discriminated kind: Text, Picture, Table, GraphicFrame, Group, Connector.</param>
/// <param name="X">Left offset in EMUs.</param>
/// <param name="Y">Top offset in EMUs.</param>
/// <param name="Width">Width in EMUs.</param>
/// <param name="Height">Height in EMUs.</param>
/// <param name="IsPlaceholder">True when this shape is a layout placeholder.</param>
/// <param name="PlaceholderType">Placeholder type string. Null for non-placeholders.</param>
/// <param name="Text">Full concatenated text content. Null for non-text shapes.</param>
/// <param name="Paragraphs">Individual paragraph strings. Null for non-text shapes.</param>
/// <param name="Table">Embedded table data. Populated only for Table shapes.</param>
/// <param name="Image">Embedded image metadata. Populated only for Picture shapes.</param>
/// <param name="Chart">Embedded chart data. Populated only for chart GraphicFrames.</param>
public record ShapeExport(
    uint? ShapeId,
    string Name,
    string ShapeType,
    long? X,
    long? Y,
    long? Width,
    long? Height,
    bool IsPlaceholder,
    string? PlaceholderType,
    string? Text,
    IReadOnlyList<string>? Paragraphs,
    TableExportData? Table = null,
    ImageExport? Image = null,
    ChartExport? Chart = null);

/// <summary>Table cell data embedded within a shape.</summary>
/// <param name="RowCount">Number of rows.</param>
/// <param name="ColumnCount">Number of columns.</param>
/// <param name="Cells">All rows of cell data. Each row is a list of cell text strings.</param>
public record TableExportData(
    int RowCount,
    int ColumnCount,
    IReadOnlyList<IReadOnlyList<string>> Cells);

/// <summary>Exported chart with series data.</summary>
/// <param name="ShapeName">Name of the chart shape.</param>
/// <param name="ChartType">Detected chart type (Column, Bar, Line, Pie, etc.).</param>
/// <param name="SeriesCount">Number of data series.</param>
/// <param name="Series">Array of series data.</param>
public record ChartExport(
    string? ShapeName,
    string? ChartType,
    int SeriesCount,
    IReadOnlyList<ChartSeriesExport> Series);

/// <summary>Single chart series data for export.</summary>
/// <param name="SeriesIndex">Zero-based series index.</param>
/// <param name="SeriesName">Display name of the series.</param>
/// <param name="Categories">Category labels.</param>
/// <param name="Values">Numeric data values.</param>
public record ChartSeriesExport(
    int SeriesIndex,
    string? SeriesName,
    IReadOnlyList<string> Categories,
    IReadOnlyList<double> Values);

/// <summary>Exported image metadata.</summary>
/// <param name="ShapeName">Name of the picture shape.</param>
/// <param name="ContentType">MIME content type (e.g. image/png).</param>
/// <param name="ImageFormat">Short format label (PNG, JPEG, etc.).</param>
/// <param name="RelationshipId">OpenXML relationship ID.</param>
/// <param name="WidthEmu">Width in EMUs.</param>
/// <param name="HeightEmu">Height in EMUs.</param>
public record ImageExport(
    string? ShapeName,
    string ContentType,
    string ImageFormat,
    string RelationshipId,
    long? WidthEmu,
    long? HeightEmu);
