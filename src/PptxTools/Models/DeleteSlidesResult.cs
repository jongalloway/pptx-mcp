namespace PptxTools.Models;

/// <summary>Structured result for Delete/DeleteAll actions of pptx_manage_slides.</summary>
/// <param name="Success">True when slide deletion completed successfully.</param>
/// <param name="DeletedCount">Number of slides deleted.</param>
/// <param name="DeletedSlideNumbers">1-based slide numbers requested for Delete. Empty for DeleteAll.</param>
/// <param name="Message">Human-readable status message.</param>
public record DeleteSlidesResult(
    bool Success,
    int DeletedCount,
    int[] DeletedSlideNumbers,
    string Message);
