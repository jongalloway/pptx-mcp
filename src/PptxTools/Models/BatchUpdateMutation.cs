namespace PptxTools.Models;

/// <summary>Text mutation request for pptx_batch_update.</summary>
/// <param name="SlideNumber">1-based slide number that contains the target shape.</param>
/// <param name="ShapeName">Exact shape name to update, ignoring case.</param>
/// <param name="NewValue">Replacement text for the shape. Newlines create separate paragraphs.</param>
public record BatchUpdateMutation(
    int SlideNumber,
    string ShapeName,
    string NewValue);
