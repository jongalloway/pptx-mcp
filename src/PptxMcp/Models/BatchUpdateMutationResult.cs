namespace PptxMcp.Models;

/// <summary>Per-mutation outcome for pptx_batch_update.</summary>
/// <param name="SlideNumber">1-based slide number that was targeted.</param>
/// <param name="ShapeName">Shape name requested by the caller.</param>
/// <param name="Success">True when the mutation was applied.</param>
/// <param name="Error">Failure message when the mutation could not be applied.</param>
/// <param name="MatchedBy">How the target was resolved, currently shapeName when successful.</param>
public record BatchUpdateMutationResult(
    int SlideNumber,
    string ShapeName,
    bool Success,
    string? Error,
    string? MatchedBy);
