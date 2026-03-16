namespace PptxMcp.Models;

/// <summary>Structured representation of a single shape on a slide.</summary>
/// <param name="ShapeId">Numeric ID assigned to the shape in the OpenXML tree.</param>
/// <param name="Name">Shape name from the drawing properties.</param>
/// <param name="ShapeType">Discriminated kind: Text, Picture, Table, GraphicFrame, Group, or Connector.</param>
/// <param name="X">Left edge offset in EMUs (English Metric Units). Null when position is inherited or unavailable.</param>
/// <param name="Y">Top edge offset in EMUs. Null when unavailable.</param>
/// <param name="Width">Width in EMUs. Null when unavailable.</param>
/// <param name="Height">Height in EMUs. Null when unavailable.</param>
/// <param name="IsPlaceholder">True when this shape is a layout placeholder.</param>
/// <param name="PlaceholderType">Placeholder type string (e.g. "title", "body", "ctrTitle"). Null for non-placeholders.</param>
/// <param name="PlaceholderIndex">Placeholder index within the layout. Null for non-placeholders.</param>
/// <param name="Text">Full concatenated text content. Paragraphs are separated by newlines. Null for non-text shapes.</param>
/// <param name="Paragraphs">Individual paragraph strings. Null for non-text shapes.</param>
/// <param name="TableRows">For Table shapes, each element is a row represented as a list of cell text strings. Null for non-table shapes.</param>
public record ShapeContent(
    uint? ShapeId,
    string Name,
    string ShapeType,
    long? X,
    long? Y,
    long? Width,
    long? Height,
    bool IsPlaceholder,
    string? PlaceholderType,
    uint? PlaceholderIndex,
    string? Text,
    IReadOnlyList<string>? Paragraphs,
    IReadOnlyList<IReadOnlyList<string>>? TableRows);
