namespace PptxTools.Models;

/// <summary>A single cell update request for pptx_update_table.</summary>
/// <param name="Row">Zero-based row index of the cell to update.</param>
/// <param name="Column">Zero-based column index of the cell to update.</param>
/// <param name="Value">New text value for the cell.</param>
public record TableCellUpdate(int Row, int Column, string Value);
