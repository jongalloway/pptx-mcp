using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxTools.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Services;

public partial class PresentationService
{
    /// <summary>
    /// Execute a heterogeneous batch of operations against a single presentation file
    /// in one open/save cycle. When <paramref name="atomic"/> is true the original file
    /// is restored if any operation fails.
    /// </summary>
    public BatchOperationResult BatchExecute(string filePath, IReadOnlyList<BatchOperation> operations, bool atomic = false)
    {
        if (operations is null || operations.Count == 0)
            return new BatchOperationResult(0, 0, 0, false, []);

        string? backupPath = null;
        if (atomic)
        {
            backupPath = filePath + ".bak";
            File.Copy(filePath, backupPath, overwrite: true);
        }

        var outcomes = new List<BatchOperationOutcome>(operations.Count);
        var modifiedSlideParts = new HashSet<SlidePart>();
        bool anyFailure = false;

        try
        {
            using var doc = PresentationDocument.Open(filePath, true);
            var slideIds = GetSlideIds(doc);

            foreach (var op in operations)
            {
                var outcome = ExecuteOperation(doc, slideIds, op, modifiedSlideParts);
                outcomes.Add(outcome);
                if (!outcome.Success)
                    anyFailure = true;
            }

            if (atomic && anyFailure)
            {
                // Don't save — close the document and restore backup below
            }
            else
            {
                foreach (var part in modifiedSlideParts)
                    part.Slide?.Save();
            }
        }
        catch (Exception ex)
        {
            // Catastrophic failure — treat as full batch failure
            if (outcomes.Count < operations.Count)
            {
                for (int i = outcomes.Count; i < operations.Count; i++)
                {
                    var op = operations[i];
                    outcomes.Add(new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false, $"Aborted: {ex.Message}", null));
                }
            }
            anyFailure = true;
        }

        bool rolledBack = false;
        if (atomic && anyFailure && backupPath is not null)
        {
            try
            {
                File.Copy(backupPath, filePath, overwrite: true);
                rolledBack = true;
            }
            catch
            {
                // Best-effort restore
            }
        }

        // Clean up backup
        if (backupPath is not null)
        {
            try { File.Delete(backupPath); } catch { }
        }

        var successCount = outcomes.Count(o => o.Success);
        return new BatchOperationResult(
            TotalOperations: outcomes.Count,
            SuccessCount: successCount,
            FailureCount: outcomes.Count - successCount,
            RolledBack: rolledBack,
            Results: outcomes);
    }

    private static BatchOperationOutcome ExecuteOperation(
        PresentationDocument doc,
        IReadOnlyList<SlideId> slideIds,
        BatchOperation op,
        HashSet<SlidePart> modifiedSlideParts)
    {
        try
        {
            return op.Type switch
            {
                BatchOperationType.UpdateText => ExecuteUpdateText(doc, slideIds, op, modifiedSlideParts),
                BatchOperationType.UpdateTableCell => ExecuteUpdateTableCell(doc, slideIds, op, modifiedSlideParts),
                BatchOperationType.UpdateShapeProperties => ExecuteUpdateShapeProperties(doc, slideIds, op, modifiedSlideParts),
                BatchOperationType.ReplaceImage => ExecuteReplaceImage(doc, slideIds, op, modifiedSlideParts),
                _ => new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false, $"Unknown operation type: {op.Type}", null)
            };
        }
        catch (Exception ex)
        {
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false, ex.Message, null);
        }
    }

    private static BatchOperationOutcome ExecuteUpdateText(
        PresentationDocument doc,
        IReadOnlyList<SlideId> slideIds,
        BatchOperation op,
        HashSet<SlidePart> modifiedSlideParts)
    {
        var result = UpdateSlideData(doc, slideIds, op.SlideNumber, op.ShapeName, null, op.NewText ?? string.Empty, out var modifiedPart);
        if (result.Success && modifiedPart is not null)
            modifiedSlideParts.Add(modifiedPart);

        return new BatchOperationOutcome(
            op.SlideNumber, op.ShapeName, op.Type,
            result.Success,
            result.Success ? null : result.Message,
            result.Success ? $"MatchedBy: {result.MatchedBy}" : null);
    }

    private static BatchOperationOutcome ExecuteUpdateTableCell(
        PresentationDocument doc,
        IReadOnlyList<SlideId> slideIds,
        BatchOperation op,
        HashSet<SlidePart> modifiedSlideParts)
    {
        if (op.TableRow is null || op.TableColumn is null)
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false, "TableRow and TableColumn are required for UpdateTableCell.", null);

        var slidePart = GetSlidePart(doc, slideIds, op.SlideNumber - 1);
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        // Find tables on the slide
        var tables = shapeTree.Elements<P.GraphicFrame>()
            .Where(gf => gf.Graphic?.GraphicData?.GetFirstChild<A.Table>() is not null)
            .ToList();

        if (tables.Count == 0)
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false, $"Slide {op.SlideNumber} has no tables.", null);

        // Resolve by shape name
        var targetFrame = tables.FirstOrDefault(gf =>
            string.Equals(
                gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value,
                op.ShapeName,
                StringComparison.OrdinalIgnoreCase));

        if (targetFrame is null)
        {
            var available = string.Join(", ",
                tables.Select(gf => gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "(unnamed)"));
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false,
                $"No table named '{op.ShapeName}' on slide {op.SlideNumber}. Available: {available}", null);
        }

        var table = targetFrame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
        var tableRows = table.Elements<A.TableRow>().ToList();
        int row = op.TableRow.Value;
        int col = op.TableColumn.Value;

        if (row < 0 || row >= tableRows.Count)
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false,
                $"Row {row} is out of range. Table has {tableRows.Count} row(s).", null);

        var cells = tableRows[row].Elements<A.TableCell>().ToList();
        if (col < 0 || col >= cells.Count)
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false,
                $"Column {col} is out of range. Row {row} has {cells.Count} column(s).", null);

        var cell = cells[col];
        UpdateTableCellText(cell, op.CellValue ?? string.Empty);
        modifiedSlideParts.Add(slidePart);

        return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, true, null,
            $"Cell [{row},{col}] updated");
    }

    private static BatchOperationOutcome ExecuteUpdateShapeProperties(
        PresentationDocument doc,
        IReadOnlyList<SlideId> slideIds,
        BatchOperation op,
        HashSet<SlidePart> modifiedSlideParts)
    {
        if (op.X is null && op.Y is null && op.Width is null && op.Height is null && op.Rotation is null)
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false, "At least one shape property (X, Y, Width, Height, Rotation) must be specified.", null);

        var slidePart = GetSlidePart(doc, slideIds, op.SlideNumber - 1);
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        // Search across all shape types that have ShapeProperties
        P.ShapeProperties? shapeProps = null;
        string? resolvedName = null;

        // Try regular shapes
        foreach (var shape in shapeTree.Elements<Shape>())
        {
            var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
            if (string.Equals(name, op.ShapeName, StringComparison.OrdinalIgnoreCase))
            {
                shapeProps = shape.ShapeProperties;
                resolvedName = name;
                break;
            }
        }

        // Try pictures
        if (shapeProps is null)
        {
            foreach (var pic in shapeTree.Elements<Picture>())
            {
                var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
                if (string.Equals(name, op.ShapeName, StringComparison.OrdinalIgnoreCase))
                {
                    shapeProps = pic.ShapeProperties;
                    resolvedName = name;
                    break;
                }
            }
        }

        // Try graphic frames (tables, charts)
        if (shapeProps is null)
        {
            foreach (var gf in shapeTree.Elements<P.GraphicFrame>())
            {
                var name = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value;
                if (string.Equals(name, op.ShapeName, StringComparison.OrdinalIgnoreCase))
                {
                    // GraphicFrame uses Transform (not ShapeProperties.Transform2D)
                    var xfrm = gf.Transform;
                    if (xfrm is null)
                    {
                        xfrm = new P.Transform();
                        gf.Transform = xfrm;
                    }
                    ApplyTransformProperties(xfrm, op);
                    modifiedSlideParts.Add(slidePart);
                    return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, true, null,
                        $"Properties updated on '{name}'");
                }
            }
        }

        if (shapeProps is null)
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false,
                $"No shape named '{op.ShapeName}' found on slide {op.SlideNumber}.", null);

        // Ensure Transform2D exists
        var transform = shapeProps.Transform2D;
        if (transform is null)
        {
            transform = new A.Transform2D();
            shapeProps.Transform2D = transform;
        }

        if (op.X is not null || op.Y is not null)
        {
            var offset = transform.Offset ?? new A.Offset();
            if (op.X is not null) offset.X = op.X.Value;
            if (op.Y is not null) offset.Y = op.Y.Value;
            transform.Offset = offset;
        }

        if (op.Width is not null || op.Height is not null)
        {
            var extents = transform.Extents ?? new A.Extents();
            if (op.Width is not null) extents.Cx = op.Width.Value;
            if (op.Height is not null) extents.Cy = op.Height.Value;
            transform.Extents = extents;
        }

        if (op.Rotation is not null)
            transform.Rotation = (int)op.Rotation.Value;

        modifiedSlideParts.Add(slidePart);
        return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, true, null,
            $"Properties updated on '{resolvedName}'");
    }

    private static BatchOperationOutcome ExecuteReplaceImage(
        PresentationDocument doc,
        IReadOnlyList<SlideId> slideIds,
        BatchOperation op,
        HashSet<SlidePart> modifiedSlideParts)
    {
        if (string.IsNullOrWhiteSpace(op.ImagePath))
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false, "ImagePath is required for ReplaceImage.", null);

        if (!File.Exists(op.ImagePath))
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false, $"Image file not found: {op.ImagePath}", null);

        var imageContentType = GetImageContentTypeString(op.ImagePath);
        if (imageContentType is null)
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false,
                $"Unsupported image format: {Path.GetExtension(op.ImagePath)}. Supported: .png, .jpg, .jpeg, .svg", null);

        var slidePart = GetSlidePart(doc, slideIds, op.SlideNumber - 1);
        var pictureTargets = GetPictureTargets(slidePart.Slide);

        var target = ResolvePictureTarget(pictureTargets, op.ShapeName, null, out var matchedBy, out var failureMessage);
        if (target is null)
            return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, false,
                failureMessage ?? $"No picture named '{op.ShapeName}' found on slide {op.SlideNumber}.", null);

        // Create new image part and feed data
        var imagePartType = GetImagePartType(op.ImagePath);
        var newImagePart = slidePart.AddImagePart(imagePartType);
        using (var stream = File.OpenRead(op.ImagePath))
            newImagePart.FeedData(stream);
        var newRelId = slidePart.GetIdOfPart(newImagePart);

        // Update blip reference
        var blipFill = target.Picture.GetFirstChild<P.BlipFill>();
        if (blipFill is null)
        {
            target.Picture.Append(new P.BlipFill(
                new A.Blip { Embed = newRelId },
                new A.Stretch(new A.FillRectangle())));
        }
        else
        {
            var blip = blipFill.GetFirstChild<A.Blip>();
            if (blip is not null)
            {
                var oldRelId = blip.Embed?.Value;
                blip.Embed = newRelId;

                // Remove SVG extension for non-SVG replacements
                var extList = blip.GetFirstChild<A.BlipExtensionList>();
                if (extList is not null && !string.Equals(imageContentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase))
                    extList.Remove();

                if (!string.IsNullOrEmpty(oldRelId) && oldRelId != newRelId)
                    TryRemoveUnusedImagePart(slidePart, oldRelId);
            }
            else
            {
                blipFill.InsertAt(new A.Blip { Embed = newRelId }, 0);
            }
        }

        modifiedSlideParts.Add(slidePart);
        return new BatchOperationOutcome(op.SlideNumber, op.ShapeName, op.Type, true, null,
            $"Image replaced in '{target.Name}' (MatchedBy: {matchedBy})");
    }

    /// <summary>Update the text content of a table cell, preserving existing run formatting.</summary>
    private static void UpdateTableCellText(A.TableCell cell, string value)
    {
        var textBody = cell.GetFirstChild<A.TextBody>();
        if (textBody is not null)
        {
            // Collapse to a single paragraph
            foreach (var p in textBody.Elements<A.Paragraph>().Skip(1).ToList())
                p.Remove();

            var firstPara = textBody.GetFirstChild<A.Paragraph>()
                ?? textBody.AppendChild(new A.Paragraph());

            // Preserve run properties from the first run
            var existingRunProps = firstPara.Elements<A.Run>()
                .FirstOrDefault()
                ?.GetFirstChild<A.RunProperties>()
                ?.CloneNode(true) as A.RunProperties;

            foreach (var r in firstPara.Elements<A.Run>().ToList())
                r.Remove();

            var newRun = new A.Run(new A.Text(value));
            if (existingRunProps is not null)
                newRun.PrependChild(existingRunProps);

            var endParaRunProps = firstPara.GetFirstChild<A.EndParagraphRunProperties>();
            if (endParaRunProps is not null)
                firstPara.InsertBefore(newRun, endParaRunProps);
            else
                firstPara.Append(newRun);
        }
        else
        {
            var newTextBody = new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(new A.Text(value)),
                    new A.EndParagraphRunProperties()));
            var tcPr = cell.GetFirstChild<A.TableCellProperties>();
            if (tcPr is not null)
                cell.InsertBefore(newTextBody, tcPr);
            else
                cell.Append(newTextBody);
        }
    }

    /// <summary>Apply position/size/rotation properties to a GraphicFrame transform.</summary>
    private static void ApplyTransformProperties(P.Transform xfrm, BatchOperation op)
    {
        if (op.X is not null || op.Y is not null)
        {
            var offset = xfrm.Offset ?? new A.Offset();
            if (op.X is not null) offset.X = op.X.Value;
            if (op.Y is not null) offset.Y = op.Y.Value;
            xfrm.Offset = offset;
        }

        if (op.Width is not null || op.Height is not null)
        {
            var extents = xfrm.Extents ?? new A.Extents();
            if (op.Width is not null) extents.Cx = op.Width.Value;
            if (op.Height is not null) extents.Cy = op.Height.Value;
            xfrm.Extents = extents;
        }

        if (op.Rotation is not null)
            xfrm.Rotation = (int)op.Rotation.Value;
    }
}
