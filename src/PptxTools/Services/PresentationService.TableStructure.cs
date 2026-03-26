using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxTools.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Services;

public partial class PresentationService
{
    public TableStructureResult AddTableRow(
        string filePath,
        int slideNumber,
        int? rowIndex = null,
        string[]? cellValues = null,
        string? tableName = null,
        int? tableIndex = null)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var (slidePart, targetFrame, table, resolvedName) = FindTableOnSlide(doc, slideNumber, tableName, tableIndex);

        var tableRows = table.Elements<A.TableRow>().ToList();
        var columnCount = table.TableGrid!.Elements<A.GridColumn>().Count();

        var insertAt = rowIndex ?? tableRows.Count; // default: append at end
        if (insertAt < 0 || insertAt > tableRows.Count)
            throw new ArgumentOutOfRangeException(nameof(rowIndex),
                $"Row index {insertAt} is out of range. Table has {tableRows.Count} row(s), valid range is 0–{tableRows.Count}.");

        var values = NormalizeCellValues(cellValues, columnCount);
        var rowHeight = tableRows.Count > 0 ? tableRows[0].Height?.Value ?? 342900L : 342900L;
        var newRow = BuildTableRow(values, rowHeight);

        if (insertAt >= tableRows.Count)
            table.Append(newRow);
        else
            table.InsertBefore(newRow, tableRows[insertAt]);

        slidePart.Slide.Save();

        var newRowCount = tableRows.Count + 1;
        return new TableStructureResult(
            Success: true,
            SlideNumber: slideNumber,
            Action: nameof(TableStructureAction.AddRow),
            TableName: resolvedName,
            RowCount: newRowCount,
            ColumnCount: columnCount,
            Message: $"Added row at index {insertAt} in table '{resolvedName}' on slide {slideNumber}. Table now has {newRowCount} rows.");
    }

    public TableStructureResult DeleteTableRow(
        string filePath,
        int slideNumber,
        int rowIndex,
        string? tableName = null,
        int? tableIndex = null)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var (slidePart, targetFrame, table, resolvedName) = FindTableOnSlide(doc, slideNumber, tableName, tableIndex);

        var tableRows = table.Elements<A.TableRow>().ToList();
        if (tableRows.Count <= 1)
            throw new InvalidOperationException("Cannot delete the last remaining row in a table.");

        if (rowIndex < 0 || rowIndex >= tableRows.Count)
            throw new ArgumentOutOfRangeException(nameof(rowIndex),
                $"Row index {rowIndex} is out of range. Table has {tableRows.Count} row(s).");

        tableRows[rowIndex].Remove();
        slidePart.Slide.Save();

        var columnCount = table.TableGrid!.Elements<A.GridColumn>().Count();
        var newRowCount = tableRows.Count - 1;
        return new TableStructureResult(
            Success: true,
            SlideNumber: slideNumber,
            Action: nameof(TableStructureAction.DeleteRow),
            TableName: resolvedName,
            RowCount: newRowCount,
            ColumnCount: columnCount,
            Message: $"Deleted row {rowIndex} from table '{resolvedName}' on slide {slideNumber}. Table now has {newRowCount} rows.");
    }

    public TableStructureResult AddTableColumn(
        string filePath,
        int slideNumber,
        int? columnIndex = null,
        string? headerText = null,
        string? tableName = null,
        int? tableIndex = null)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var (slidePart, targetFrame, table, resolvedName) = FindTableOnSlide(doc, slideNumber, tableName, tableIndex);

        var gridColumns = table.TableGrid!.Elements<A.GridColumn>().ToList();
        var currentColumnCount = gridColumns.Count;
        var insertAt = columnIndex ?? currentColumnCount; // default: append at end

        if (insertAt < 0 || insertAt > currentColumnCount)
            throw new ArgumentOutOfRangeException(nameof(columnIndex),
                $"Column index {insertAt} is out of range. Table has {currentColumnCount} column(s), valid range is 0–{currentColumnCount}.");

        // Calculate column width: use average of existing columns
        var columnWidth = currentColumnCount > 0
            ? gridColumns.Average(gc => gc.Width?.Value ?? 914400L)
            : 914400L;

        // 1. Update TableGrid
        var newGridColumn = new A.GridColumn { Width = (long)columnWidth };
        if (insertAt >= currentColumnCount)
            table.TableGrid.Append(newGridColumn);
        else
            table.TableGrid.InsertBefore(newGridColumn, gridColumns[insertAt]);

        // 2. Update every row: insert a new cell at the same column position
        var tableRows = table.Elements<A.TableRow>().ToList();
        for (int rowIdx = 0; rowIdx < tableRows.Count; rowIdx++)
        {
            var cellText = (rowIdx == 0 && headerText is not null) ? headerText : string.Empty;
            var newCell = BuildTableCell(cellText);

            var existingCells = tableRows[rowIdx].Elements<A.TableCell>().ToList();
            if (insertAt >= existingCells.Count)
                tableRows[rowIdx].Append(newCell);
            else
                tableRows[rowIdx].InsertBefore(newCell, existingCells[insertAt]);
        }

        slidePart.Slide.Save();

        var newColumnCount = currentColumnCount + 1;
        var rowCount = tableRows.Count;
        return new TableStructureResult(
            Success: true,
            SlideNumber: slideNumber,
            Action: nameof(TableStructureAction.AddColumn),
            TableName: resolvedName,
            RowCount: rowCount,
            ColumnCount: newColumnCount,
            Message: $"Added column at index {insertAt} in table '{resolvedName}' on slide {slideNumber}. Table now has {newColumnCount} columns.");
    }

    public TableStructureResult DeleteTableColumn(
        string filePath,
        int slideNumber,
        int columnIndex,
        string? tableName = null,
        int? tableIndex = null)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var (slidePart, targetFrame, table, resolvedName) = FindTableOnSlide(doc, slideNumber, tableName, tableIndex);

        var gridColumns = table.TableGrid!.Elements<A.GridColumn>().ToList();
        var currentColumnCount = gridColumns.Count;

        if (currentColumnCount <= 1)
            throw new InvalidOperationException("Cannot delete the last remaining column in a table.");

        if (columnIndex < 0 || columnIndex >= currentColumnCount)
            throw new ArgumentOutOfRangeException(nameof(columnIndex),
                $"Column index {columnIndex} is out of range. Table has {currentColumnCount} column(s).");

        // 1. Remove GridColumn
        gridColumns[columnIndex].Remove();

        // 2. Remove corresponding cell from every row
        var tableRows = table.Elements<A.TableRow>().ToList();
        foreach (var row in tableRows)
        {
            var cells = row.Elements<A.TableCell>().ToList();
            if (columnIndex < cells.Count)
                cells[columnIndex].Remove();
        }

        slidePart.Slide.Save();

        var newColumnCount = currentColumnCount - 1;
        var rowCount = tableRows.Count;
        return new TableStructureResult(
            Success: true,
            SlideNumber: slideNumber,
            Action: nameof(TableStructureAction.DeleteColumn),
            TableName: resolvedName,
            RowCount: rowCount,
            ColumnCount: newColumnCount,
            Message: $"Deleted column {columnIndex} from table '{resolvedName}' on slide {slideNumber}. Table now has {newColumnCount} columns.");
    }

    public TableStructureResult MergeTableCells(
        string filePath,
        int slideNumber,
        int startRow,
        int startCol,
        int endRow,
        int endCol,
        string? tableName = null,
        int? tableIndex = null)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var (slidePart, targetFrame, table, resolvedName) = FindTableOnSlide(doc, slideNumber, tableName, tableIndex);

        var tableRows = table.Elements<A.TableRow>().ToList();
        var columnCount = table.TableGrid!.Elements<A.GridColumn>().Count();

        if (startRow < 0 || startRow >= tableRows.Count)
            throw new ArgumentOutOfRangeException(nameof(startRow), $"startRow {startRow} is out of range. Table has {tableRows.Count} row(s).");
        if (endRow < startRow || endRow >= tableRows.Count)
            throw new ArgumentOutOfRangeException(nameof(endRow), $"endRow {endRow} is out of range. Must be >= startRow ({startRow}) and < {tableRows.Count}.");
        if (startCol < 0 || startCol >= columnCount)
            throw new ArgumentOutOfRangeException(nameof(startCol), $"startCol {startCol} is out of range. Table has {columnCount} column(s).");
        if (endCol < startCol || endCol >= columnCount)
            throw new ArgumentOutOfRangeException(nameof(endCol), $"endCol {endCol} is out of range. Must be >= startCol ({startCol}) and < {columnCount}.");

        var horizontalSpan = endCol - startCol + 1;
        var verticalSpan = endRow - startRow + 1;

        for (int r = startRow; r <= endRow; r++)
        {
            var cells = tableRows[r].Elements<A.TableCell>().ToList();

            for (int c = startCol; c <= endCol; c++)
            {
                if (c >= cells.Count) continue;
                var cell = cells[c];

                if (r == startRow && c == startCol)
                {
                    // Anchor cell: set GridSpan and RowSpan on the cell element
                    if (horizontalSpan > 1)
                        cell.GridSpan = horizontalSpan;
                    if (verticalSpan > 1)
                        cell.RowSpan = verticalSpan;
                }
                else
                {
                    // Non-anchor cells: mark as merged on the cell element
                    if (horizontalSpan > 1 && c > startCol)
                        cell.HorizontalMerge = true;
                    if (verticalSpan > 1 && r > startRow)
                        cell.VerticalMerge = true;
                }
            }
        }

        slidePart.Slide.Save();

        return new TableStructureResult(
            Success: true,
            SlideNumber: slideNumber,
            Action: nameof(TableStructureAction.MergeCells),
            TableName: resolvedName,
            RowCount: tableRows.Count,
            ColumnCount: columnCount,
            Message: $"Merged cells [{startRow},{startCol}] to [{endRow},{endCol}] in table '{resolvedName}' on slide {slideNumber}.");
    }

    /// <summary>Find a table on a slide by name or index, reusing the UpdateTable lookup pattern.</summary>
    private static (SlidePart slidePart, P.GraphicFrame frame, A.Table table, string? resolvedName) FindTableOnSlide(
        PresentationDocument doc,
        int slideNumber,
        string? tableName,
        int? tableIndex)
    {
        var slideIds = GetSlideIds(doc);
        var slidePart = GetSlidePart(doc, slideIds, slideNumber - 1);
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        var tables = shapeTree.Elements<P.GraphicFrame>()
            .Where(gf => gf.Graphic?.GraphicData?.GetFirstChild<A.Table>() is not null)
            .ToList();

        if (tables.Count == 0)
            throw new InvalidOperationException($"Slide {slideNumber} has no tables.");

        P.GraphicFrame? targetFrame = null;

        if (!string.IsNullOrWhiteSpace(tableName))
        {
            targetFrame = tables.FirstOrDefault(gf =>
                string.Equals(
                    gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value,
                    tableName,
                    StringComparison.OrdinalIgnoreCase));

            if (targetFrame is null)
            {
                var available = string.Join(", ",
                    tables.Select((gf, i) =>
                        $"{i}:{gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "(unnamed)"}"));
                throw new InvalidOperationException(
                    $"No table named '{tableName}' found on slide {slideNumber}. Available tables: {available}");
            }
        }
        else if (tableIndex.HasValue)
        {
            if (tableIndex.Value < 0 || tableIndex.Value >= tables.Count)
                throw new ArgumentOutOfRangeException(nameof(tableIndex),
                    $"Table index {tableIndex.Value} is out of range. Slide {slideNumber} has {tables.Count} table(s).");
            targetFrame = tables[tableIndex.Value];
        }
        else
        {
            if (tables.Count > 1)
                throw new InvalidOperationException(
                    $"Slide {slideNumber} has {tables.Count} tables. Specify tableName or tableIndex to identify the target.");
            targetFrame = tables[0];
        }

        var table = targetFrame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
        var resolvedName = targetFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value;

        return (slidePart, targetFrame, table, resolvedName);
    }

    /// <summary>Normalize cell values to match the expected column count, padding with empty strings.</summary>
    private static string[] NormalizeCellValues(string[]? values, int columnCount)
    {
        var result = new string[columnCount];
        for (int i = 0; i < columnCount; i++)
            result[i] = (values is not null && i < values.Length) ? (values[i] ?? string.Empty) : string.Empty;
        return result;
    }

    /// <summary>Build a single table cell with proper TextBody structure.</summary>
    private static A.TableCell BuildTableCell(string text) =>
        new(
            new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(new A.Text(text ?? string.Empty)),
                    new A.EndParagraphRunProperties())),
            new A.TableCellProperties());
}
