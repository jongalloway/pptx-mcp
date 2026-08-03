namespace PptxTools.Models;

/// <summary>Actions for the consolidated pptx_table_structure tool.</summary>
public enum TableStructureAction
{
    /// <summary>Add a new row to the table at a specified position (default: end).</summary>
    AddRow,

    /// <summary>Delete an existing row from the table by index.</summary>
    DeleteRow,

    /// <summary>Add a new column to the table at a specified position (default: end).</summary>
    AddColumn,

    /// <summary>Delete an existing column from the table by index.</summary>
    DeleteColumn,

    /// <summary>Merge a rectangular range of cells in the table.</summary>
    MergeCells
}

/// <summary>Structured result for pptx_table_structure operations.</summary>
/// <param name="Success">True when the structural change was applied successfully.</param>
/// <param name="SlideNumber">1-based slide number containing the table.</param>
/// <param name="Action">The structural action that was performed.</param>
/// <param name="TableName">Name of the matched table's GraphicFrame.</param>
/// <param name="RowCount">Total row count after the operation.</param>
/// <param name="ColumnCount">Total column count after the operation.</param>
/// <param name="Message">Human-readable status message describing the outcome.</param>
public record TableStructureResult(
    bool Success,
    int SlideNumber,
    string Action,
    string? TableName,
    int RowCount,
    int ColumnCount,
    string Message);
