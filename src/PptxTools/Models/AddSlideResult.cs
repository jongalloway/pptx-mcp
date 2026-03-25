namespace PptxTools.Models;

/// <summary>Structured result for the Add action of pptx_manage_slides.</summary>
/// <param name="Success">True when the slide was created successfully.</param>
/// <param name="SlideNumber">1-based slide number of the created slide.</param>
/// <param name="LayoutName">Layout name used (null if default was used).</param>
/// <param name="Message">Human-readable status message.</param>
public record AddSlideResult(
    bool Success,
    int? SlideNumber,
    string? LayoutName,
    string Message);
