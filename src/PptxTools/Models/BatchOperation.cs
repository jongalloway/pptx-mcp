namespace PptxTools.Models;

/// <summary>Discriminator for the type of operation in a batch execute request.</summary>
public enum BatchOperationType
{
    /// <summary>Replace text in a named shape, preserving formatting.</summary>
    UpdateText,
    /// <summary>Update a specific cell in a named table.</summary>
    UpdateTableCell,
    /// <summary>Modify position, size, or rotation of a shape.</summary>
    UpdateShapeProperties,
    /// <summary>Swap the image source of a named picture shape.</summary>
    ReplaceImage
}

/// <summary>
/// A single operation in a pptx_batch_execute request.
/// Fields are populated based on <see cref="Type"/>; unused fields should be null.
/// </summary>
/// <param name="SlideNumber">1-based slide number containing the target shape.</param>
/// <param name="ShapeName">Exact shape/table/picture name to match, ignoring case.</param>
/// <param name="Type">The kind of operation to perform.</param>
/// <param name="NewText">Replacement text (UpdateText).</param>
/// <param name="TableRow">0-based row index (UpdateTableCell).</param>
/// <param name="TableColumn">0-based column index (UpdateTableCell).</param>
/// <param name="CellValue">New cell text (UpdateTableCell).</param>
/// <param name="X">Horizontal offset in EMUs (UpdateShapeProperties).</param>
/// <param name="Y">Vertical offset in EMUs (UpdateShapeProperties).</param>
/// <param name="Width">Width in EMUs (UpdateShapeProperties).</param>
/// <param name="Height">Height in EMUs (UpdateShapeProperties).</param>
/// <param name="Rotation">Rotation in 60,000ths of a degree (UpdateShapeProperties).</param>
/// <param name="ImagePath">Path to replacement image file (ReplaceImage).</param>
public record BatchOperation(
    int SlideNumber,
    string ShapeName,
    BatchOperationType Type,
    // UpdateText
    string? NewText = null,
    // UpdateTableCell
    int? TableRow = null,
    int? TableColumn = null,
    string? CellValue = null,
    // UpdateShapeProperties
    long? X = null,
    long? Y = null,
    long? Width = null,
    long? Height = null,
    long? Rotation = null,
    // ReplaceImage
    string? ImagePath = null);
