namespace PptxMcp.Models;

/// <summary>Structured result for pptx_batch_update.</summary>
/// <param name="TotalMutations">Number of requested mutations.</param>
/// <param name="SuccessCount">Number of mutations applied successfully.</param>
/// <param name="FailureCount">Number of mutations that failed.</param>
/// <param name="Results">Per-mutation outcomes in request order.</param>
public record BatchUpdateResult(
    int TotalMutations,
    int SuccessCount,
    int FailureCount,
    IReadOnlyList<BatchUpdateMutationResult> Results);
