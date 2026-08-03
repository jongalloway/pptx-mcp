namespace PptxTools.Models;

/// <summary>Structured result for pptx_update_table.</summary>
/// <param name="Success">True when the table updates were applied successfully.</param>
/// <param name="SlideNumber">1-based slide number where the table was updated.</param>
/// <param name="TableName">Name of the matched table's GraphicFrame.</param>
/// <param name="MatchedBy">How the table was located: tableName or tableIndex.</param>
/// <param name="CellsUpdated">Number of cells that were updated.</param>
/// <param name="CellsSkipped">Number of update requests that were skipped (out of range).</param>
/// <param name="Message">Human-readable status message describing the outcome.</param>
public record TableUpdateResult(
    bool Success,
    int SlideNumber,
    string? TableName,
    string? MatchedBy,
    int CellsUpdated,
    int CellsSkipped,
    string Message);
