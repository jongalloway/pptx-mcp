namespace PptxTools.Models;

/// <summary>Structured result for slide text updates performed by pptx_update_slide_data.</summary>
/// <param name="Success">True when the requested update was applied.</param>
/// <param name="SlideNumber">1-based slide number that was targeted.</param>
/// <param name="RequestedShapeName">Shape name requested by the caller, when provided.</param>
/// <param name="RequestedPlaceholderIndex">Zero-based fallback index requested by the caller, when provided.</param>
/// <param name="MatchedBy">How the target shape was resolved: shapeName, placeholderIndex, or placeholderIndexFallback.</param>
/// <param name="ResolvedShapeName">Resolved shape name from the slide.</param>
/// <param name="ResolvedShapeIndex">Zero-based index of the resolved text-capable shape on the slide.</param>
/// <param name="ResolvedShapeId">OpenXML shape ID of the resolved shape.</param>
/// <param name="PlaceholderType">Placeholder type when the resolved shape is a placeholder.</param>
/// <param name="LayoutPlaceholderIndex">Layout placeholder index when present in the slide markup.</param>
/// <param name="PreviousText">Text content before the update, if the target shape was resolved.</param>
/// <param name="NewText">Replacement text that was requested.</param>
/// <param name="Message">Human-readable status message describing the outcome.</param>
public record SlideDataUpdateResult(
    bool Success,
    int SlideNumber,
    string? RequestedShapeName,
    int? RequestedPlaceholderIndex,
    string? MatchedBy,
    string? ResolvedShapeName,
    int? ResolvedShapeIndex,
    uint? ResolvedShapeId,
    string? PlaceholderType,
    uint? LayoutPlaceholderIndex,
    string? PreviousText,
    string NewText,
    string Message);
