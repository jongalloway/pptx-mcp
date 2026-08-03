namespace PptxTools.Models;

/// <summary>Structured result for pptx_duplicate_slide.</summary>
/// <param name="Success">True when the slide was duplicated successfully.</param>
/// <param name="NewSlideNumber">1-based slide number of the duplicated slide when successful.</param>
/// <param name="ShapesCopied">Number of shapes copied onto the duplicated slide.</param>
/// <param name="OverridesApplied">Number of placeholder override values applied to the duplicate.</param>
/// <param name="Message">Human-readable status message describing the outcome.</param>
public record DuplicateSlideResult(
    bool Success,
    int? NewSlideNumber,
    int ShapesCopied,
    int OverridesApplied,
    string Message);
