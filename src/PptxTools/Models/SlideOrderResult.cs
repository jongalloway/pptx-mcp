namespace PptxTools.Models;

/// <summary>Structured result for pptx_reorder_slides actions.</summary>
/// <param name="Success">True when the operation completed successfully.</param>
/// <param name="Action">The action that was performed (Move or Reorder).</param>
/// <param name="Message">Human-readable status message.</param>
public record SlideOrderResult(
    bool Success,
    string Action,
    string Message);
