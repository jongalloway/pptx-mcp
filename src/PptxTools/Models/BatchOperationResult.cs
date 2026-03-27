namespace PptxTools.Models;

/// <summary>Structured result for pptx_batch_execute.</summary>
/// <param name="TotalOperations">Number of requested operations.</param>
/// <param name="SuccessCount">Number of operations applied successfully.</param>
/// <param name="FailureCount">Number of operations that failed.</param>
/// <param name="RolledBack">True when atomic mode was requested and the file was restored after a failure.</param>
/// <param name="Results">Per-operation outcomes in request order.</param>
public record BatchOperationResult(
    int TotalOperations,
    int SuccessCount,
    int FailureCount,
    bool RolledBack,
    IReadOnlyList<BatchOperationOutcome> Results);

/// <summary>Per-operation outcome for pptx_batch_execute.</summary>
/// <param name="SlideNumber">1-based slide number that was targeted.</param>
/// <param name="ShapeName">Shape name requested by the caller.</param>
/// <param name="Type">The operation type that was attempted.</param>
/// <param name="Success">True when the operation was applied.</param>
/// <param name="Error">Failure message when the operation could not be applied.</param>
/// <param name="Detail">Human-readable detail such as resolution method or cell coordinates.</param>
public record BatchOperationOutcome(
    int SlideNumber,
    string ShapeName,
    BatchOperationType Type,
    bool Success,
    string? Error,
    string? Detail);
