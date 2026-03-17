namespace PptxMcp.Models;

/// <summary>Structured result for pptx_add_slide_from_layout.</summary>
/// <param name="Success">True when the slide was created successfully.</param>
/// <param name="SlideNumber">1-based slide number of the created slide when successful.</param>
/// <param name="LayoutName">Resolved layout name used to create the slide.</param>
/// <param name="PlaceholdersPopulated">Number of placeholders populated from the request.</param>
/// <param name="Message">Human-readable status message describing the outcome.</param>
public record AddSlideFromLayoutResult(
    bool Success,
    int? SlideNumber,
    string? LayoutName,
    int PlaceholdersPopulated,
    string Message);
