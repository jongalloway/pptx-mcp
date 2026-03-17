namespace PptxMcp.Models;

/// <summary>Structured result for pptx_insert_table.</summary>
/// <param name="Success">True when the table was inserted successfully.</param>
/// <param name="SlideNumber">1-based slide number where the table was inserted.</param>
/// <param name="TableName">Name assigned to the table's GraphicFrame.</param>
/// <param name="TableShapeId">OpenXML shape ID assigned to the table.</param>
/// <param name="TableIndex">Zero-based index of this table among all tables on the slide.</param>
/// <param name="RowCount">Number of rows in the inserted table (including header row).</param>
/// <param name="ColumnCount">Number of columns in the inserted table.</param>
/// <param name="Message">Human-readable status message describing the outcome.</param>
public record TableInsertResult(
    bool Success,
    int SlideNumber,
    string? TableName,
    uint? TableShapeId,
    int? TableIndex,
    int RowCount,
    int ColumnCount,
    string Message);
