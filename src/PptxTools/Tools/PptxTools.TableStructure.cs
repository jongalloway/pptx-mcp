using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxTools.Models;

namespace PptxTools.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Perform structural changes on an existing table in a PowerPoint presentation.
    /// Available actions:
    /// - AddRow: Insert a new row at a position (default: end). Optionally provide cell values.
    /// - DeleteRow: Remove a row by its zero-based index.
    /// - AddColumn: Insert a new column at a position (default: end). Optionally provide header text.
    /// - DeleteColumn: Remove a column by its zero-based index.
    /// - MergeCells: Merge a rectangular range of cells specified by startRow/startCol/endRow/endCol (all zero-based).
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideNumber">1-based slide number containing the table.</param>
    /// <param name="action">The structural operation to perform: AddRow, DeleteRow, AddColumn, DeleteColumn, or MergeCells.</param>
    /// <param name="tableName">Optional table name to match (case-insensitive). Takes precedence over tableIndex.</param>
    /// <param name="tableIndex">Optional zero-based index among tables on the slide. Used when tableName is not provided.</param>
    /// <param name="rowIndex">Zero-based row index. Required for DeleteRow. Optional for AddRow (defaults to end).</param>
    /// <param name="columnIndex">Zero-based column index. Required for DeleteColumn. Optional for AddColumn (defaults to end).</param>
    /// <param name="cellValues">Optional array of cell text values for AddRow. Padded or truncated to match column count.</param>
    /// <param name="headerText">Optional header text for the new column in AddColumn (applied to first row only).</param>
    /// <param name="startRow">Zero-based starting row for MergeCells.</param>
    /// <param name="startCol">Zero-based starting column for MergeCells.</param>
    /// <param name="endRow">Zero-based ending row for MergeCells (inclusive).</param>
    /// <param name="endCol">Zero-based ending column for MergeCells (inclusive).</param>
    [McpServerTool(Title = "Table Structure")]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["AddRow","DeleteRow","AddColumn","DeleteColumn","MergeCells"]""")]
    public partial Task<string> pptx_table_structure(
        string filePath,
        int slideNumber,
        TableStructureAction action,
        string? tableName = null,
        int? tableIndex = null,
        int? rowIndex = null,
        int? columnIndex = null,
        string[]? cellValues = null,
        string? headerText = null,
        int? startRow = null,
        int? startCol = null,
        int? endRow = null,
        int? endCol = null)
    {
        TableStructureResult makeError(string message) => new(
            Success: false,
            SlideNumber: slideNumber,
            Action: action.ToString(),
            TableName: tableName,
            RowCount: 0,
            ColumnCount: 0,
            Message: message);

        return action switch
        {
            TableStructureAction.AddRow => ExecuteToolStructured(filePath,
                () => _service.AddTableRow(filePath, slideNumber, rowIndex, cellValues, tableName, tableIndex),
                makeError),

            TableStructureAction.DeleteRow => ExecuteToolStructured(filePath,
                () =>
                {
                    if (rowIndex is null)
                        throw new ArgumentException("rowIndex is required for the DeleteRow action.");
                    return _service.DeleteTableRow(filePath, slideNumber, rowIndex.Value, tableName, tableIndex);
                },
                makeError),

            TableStructureAction.AddColumn => ExecuteToolStructured(filePath,
                () => _service.AddTableColumn(filePath, slideNumber, columnIndex, headerText, tableName, tableIndex),
                makeError),

            TableStructureAction.DeleteColumn => ExecuteToolStructured(filePath,
                () =>
                {
                    if (columnIndex is null)
                        throw new ArgumentException("columnIndex is required for the DeleteColumn action.");
                    return _service.DeleteTableColumn(filePath, slideNumber, columnIndex.Value, tableName, tableIndex);
                },
                makeError),

            TableStructureAction.MergeCells => ExecuteToolStructured(filePath,
                () =>
                {
                    if (startRow is null || startCol is null || endRow is null || endCol is null)
                        throw new ArgumentException("startRow, startCol, endRow, and endCol are all required for the MergeCells action.");
                    return _service.MergeTableCells(filePath, slideNumber, startRow.Value, startCol.Value, endRow.Value, endCol.Value, tableName, tableIndex);
                },
                makeError),

            _ => Task.FromResult(JsonSerializer.Serialize(
                makeError($"Unknown action: {action}. Valid actions: AddRow, DeleteRow, AddColumn, DeleteColumn, MergeCells."),
                IndentedJson))
        };
    }
}
