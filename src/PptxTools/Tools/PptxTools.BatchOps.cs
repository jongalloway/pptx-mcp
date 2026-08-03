using System.Text.Json;
using ModelContextProtocol.Server;
using PptxTools.Models;

namespace PptxTools.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Execute a batch of mixed operations against a presentation in a single open/save cycle.
    /// Supports text updates, table cell updates, shape property changes, and image replacements.
    /// When atomic is true, the file is restored to its original state if any operation fails.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="operations">Array of operations. Each must include slideNumber, shapeName, and type. Additional fields depend on the operation type.</param>
    /// <param name="atomic">When true, all changes are rolled back if any operation fails. Defaults to false.</param>
    [McpServerTool(Title = "Batch Execute")]
    public partial Task<string> pptx_batch_execute(
        string filePath,
        BatchOperation[] operations,
        bool atomic = false)
    {
        var requestedOps = operations ?? [];
        if (requestedOps.Length == 0)
            return Task.FromResult(JsonSerializer.Serialize(
                new BatchOperationResult(0, 0, 0, false, []), IndentedJson));

        return ExecuteToolStructured(filePath,
            () => _service.BatchExecute(filePath, requestedOps, atomic),
            error => new BatchOperationResult(
                TotalOperations: requestedOps.Length,
                SuccessCount: 0,
                FailureCount: requestedOps.Length,
                RolledBack: atomic,
                Results: requestedOps
                    .Select(op => new BatchOperationOutcome(
                        SlideNumber: op.SlideNumber,
                        ShapeName: op.ShapeName,
                        Type: op.Type,
                        Success: false,
                        Error: error,
                        Detail: null))
                    .ToArray()));
    }
}
