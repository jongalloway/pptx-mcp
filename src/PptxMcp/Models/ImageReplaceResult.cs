namespace PptxMcp.Models;

/// <summary>Structured result for image replacement performed by pptx_replace_image.</summary>
/// <param name="Success">True when the image was replaced successfully.</param>
/// <param name="SlideNumber">1-based slide number that was targeted.</param>
/// <param name="ShapeName">Resolved shape name of the picture that was updated.</param>
/// <param name="MatchedBy">How the target picture was resolved: shapeName or shapeIndex.</param>
/// <param name="PreviousImageContentType">Content type of the image that was replaced, if available.</param>
/// <param name="NewImageContentType">Content type of the replacement image.</param>
/// <param name="AltText">Alt text applied to the picture, if any.</param>
/// <param name="Message">Human-readable status message describing the outcome.</param>
public record ImageReplaceResult(
    bool Success,
    int SlideNumber,
    string? ShapeName,
    string? MatchedBy,
    string? PreviousImageContentType,
    string? NewImageContentType,
    string? AltText,
    string Message);
