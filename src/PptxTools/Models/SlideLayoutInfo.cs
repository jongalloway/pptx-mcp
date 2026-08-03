namespace PptxTools.Models;

/// <summary>Metadata about a single slide layout within a presentation.</summary>
/// <param name="Index">Zero-based index of this layout across all masters.</param>
/// <param name="Name">Display name of the layout (from <c>cSld/@name</c>).</param>
/// <param name="LayoutType">OOXML semantic type string (e.g. "title", "obj", "blank"), or <c>null</c> when not set.</param>
/// <param name="MasterName">Display name of the parent slide master.</param>
/// <param name="PlaceholderCount">Number of placeholder shapes on this layout.</param>
/// <param name="PlaceholderTypes">List of placeholder type strings (e.g. "title", "body", "pic"); typeless body placeholders appear as "body".</param>
/// <param name="NonPlaceholderShapeCount">Number of shapes that are not placeholders.</param>
/// <param name="TotalShapeCount">Total number of shapes on this layout (placeholders + non-placeholders).</param>
/// <param name="HasTitlePlaceholder">True when the layout contains a title or centered-title placeholder.</param>
/// <param name="HasBodyPlaceholder">True when the layout contains a body, subtitle, or typeless-indexed placeholder.</param>
/// <param name="HasPicturePlaceholder">True when the layout contains a picture placeholder.</param>
public record SlideLayoutInfo(
    int Index,
    string Name,
    string? LayoutType,
    string MasterName,
    int PlaceholderCount,
    IReadOnlyList<string> PlaceholderTypes,
    int NonPlaceholderShapeCount,
    int TotalShapeCount,
    bool HasTitlePlaceholder,
    bool HasBodyPlaceholder,
    bool HasPicturePlaceholder);
