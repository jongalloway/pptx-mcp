using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PptxTools.Models;

namespace PptxTools.Services;

public partial class PresentationService
{
    private static readonly IReadOnlyDictionary<string, PlaceholderValues> PlaceholderTypeMap =
        new Dictionary<string, PlaceholderValues>(StringComparer.OrdinalIgnoreCase)
        {
            ["title"] = PlaceholderValues.Title,
            ["ctrTitle"] = PlaceholderValues.CenteredTitle,
            ["subTitle"] = PlaceholderValues.SubTitle,
            ["body"] = PlaceholderValues.Body,
            ["pic"] = PlaceholderValues.Picture,
            ["obj"] = PlaceholderValues.Object,
            ["chart"] = PlaceholderValues.Chart,
            ["ftr"] = PlaceholderValues.Footer,
            ["dt"] = PlaceholderValues.DateAndTime,
            ["sldNum"] = PlaceholderValues.SlideNumber,
            ["tbl"] = PlaceholderValues.Table,
            ["clipArt"] = PlaceholderValues.ClipArt,
            ["dgm"] = PlaceholderValues.Diagram,
            ["media"] = PlaceholderValues.Media,
            ["hdr"] = PlaceholderValues.Header
        };

    public ModifyLayoutPlaceholderResult ModifyLayoutPlaceholder(string filePath, string layoutName, int placeholderIndex, string newType, int? newIdx)
    {
        if (placeholderIndex < 0)
        {
            return new ModifyLayoutPlaceholderResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                PlaceholderIndex: placeholderIndex,
                NewType: newType,
                NewIdx: newIdx,
                Message: "placeholderIndex must be zero or greater.");
        }

        if (newIdx is < 0)
        {
            return new ModifyLayoutPlaceholderResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                PlaceholderIndex: placeholderIndex,
                NewType: newType,
                NewIdx: newIdx,
                Message: "newIdx must be zero or greater when provided.");
        }

        if (!TryResolvePlaceholderType(newType, out var newTypeValue, out var error))
        {
            return new ModifyLayoutPlaceholderResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                PlaceholderIndex: placeholderIndex,
                NewType: newType,
                NewIdx: newIdx,
                Message: error!);
        }

        using var doc = PresentationDocument.Open(filePath, isEditable: true);
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation part is missing.");

        var targetPart = FindLayoutPart(presentationPart, layoutName);
        if (targetPart is null)
        {
            return new ModifyLayoutPlaceholderResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                PlaceholderIndex: placeholderIndex,
                NewType: newType,
                NewIdx: newIdx,
                Message: $"Layout '{layoutName}' was not found in the presentation.");
        }

        var shapeTree = targetPart.SlideLayout?.CommonSlideData?.ShapeTree;
        if (shapeTree is null)
        {
            return new ModifyLayoutPlaceholderResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                PlaceholderIndex: placeholderIndex,
                NewType: newType,
                NewIdx: newIdx,
                Message: $"Layout '{layoutName}' has no shape tree.");
        }

        var placeholders = shapeTree.ChildElements
            .Select(GetLayoutPlaceholderShape)
            .Where(ph => ph is not null)
            .Cast<PlaceholderShape>()
            .ToList();

        if (placeholderIndex >= placeholders.Count)
        {
            return new ModifyLayoutPlaceholderResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                PlaceholderIndex: placeholderIndex,
                NewType: newType,
                NewIdx: newIdx,
                Message: $"placeholderIndex {placeholderIndex} is out of range for layout '{layoutName}'. Found {placeholders.Count} placeholder(s).");
        }

        var placeholder = placeholders[placeholderIndex];
        placeholder.Type = newTypeValue;
        if (newIdx.HasValue)
            placeholder.Index = (uint)newIdx.Value;

        targetPart.SlideLayout!.Save();

        return new ModifyLayoutPlaceholderResult(
            Success: true,
            FilePath: filePath,
            LayoutName: layoutName,
            PlaceholderIndex: placeholderIndex,
            NewType: newType,
            NewIdx: newIdx,
            Message: $"Updated placeholder {placeholderIndex} on layout '{layoutName}' to type '{newType}'." + (newIdx.HasValue ? $" Set idx={newIdx.Value}." : string.Empty));
    }

    public AddLayoutPlaceholderResult AddLayoutPlaceholder(string filePath, string layoutName, string type, int idx, long x, long y, long cx, long cy)
    {
        if (idx < 0)
        {
            return new AddLayoutPlaceholderResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                Type: type,
                Idx: idx,
                X: x,
                Y: y,
                Cx: cx,
                Cy: cy,
                Message: "idx must be zero or greater.");
        }

        if (!TryResolvePlaceholderType(type, out var placeholderTypeValue, out var error))
        {
            return new AddLayoutPlaceholderResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                Type: type,
                Idx: idx,
                X: x,
                Y: y,
                Cx: cx,
                Cy: cy,
                Message: error!);
        }

        ValidationHelpers.ValidateEmuValue(x, nameof(x));
        ValidationHelpers.ValidateEmuValue(y, nameof(y));
        ValidationHelpers.ValidateEmuValue(cx, nameof(cx));
        ValidationHelpers.ValidateEmuValue(cy, nameof(cy));

        using var doc = PresentationDocument.Open(filePath, isEditable: true);
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation part is missing.");

        var targetPart = FindLayoutPart(presentationPart, layoutName);
        if (targetPart is null)
        {
            return new AddLayoutPlaceholderResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                Type: type,
                Idx: idx,
                X: x,
                Y: y,
                Cx: cx,
                Cy: cy,
                Message: $"Layout '{layoutName}' was not found in the presentation.");
        }

        var shapeTree = targetPart.SlideLayout?.CommonSlideData?.ShapeTree;
        if (shapeTree is null)
        {
            return new AddLayoutPlaceholderResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                Type: type,
                Idx: idx,
                X: x,
                Y: y,
                Cx: cx,
                Cy: cy,
                Message: $"Layout '{layoutName}' has no shape tree.");
        }

        var newShapeId = GetMaxShapeId(shapeTree) + 1;
        var placeholderShape = new Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = newShapeId, Name = $"Placeholder {newShapeId}" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape
                    {
                        Type = placeholderTypeValue,
                        Index = (uint)idx
                    })),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.EndParagraphRunProperties())));

        shapeTree.Append(placeholderShape);
        targetPart.SlideLayout!.Save();

        return new AddLayoutPlaceholderResult(
            Success: true,
            FilePath: filePath,
            LayoutName: layoutName,
            Type: type,
            Idx: idx,
            X: x,
            Y: y,
            Cx: cx,
            Cy: cy,
            Message: $"Added placeholder type '{type}' (idx={idx}) to layout '{layoutName}'.");
    }

    private static SlideLayoutPart? FindLayoutPart(PresentationPart presentationPart, string layoutName)
    {
        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            foreach (var layoutPart in masterPart.SlideLayoutParts)
            {
                var name = layoutPart.SlideLayout?.CommonSlideData?.Name?.Value ?? string.Empty;
                if (string.Equals(name, layoutName, StringComparison.OrdinalIgnoreCase))
                    return layoutPart;
            }
        }

        return null;
    }

    private static bool TryResolvePlaceholderType(string type, out PlaceholderValues value, out string? error)
    {
        if (PlaceholderTypeMap.TryGetValue(type, out value))
        {
            error = null;
            return true;
        }

        value = default;
        var valid = string.Join(", ", PlaceholderTypeMap.Keys.Order());
        error = $"Unknown placeholder type '{type}'. Valid values: {valid}.";
        return false;
    }

    private static PlaceholderShape? GetLayoutPlaceholderShape(OpenXmlElement element) => element switch
    {
        Shape shape => shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
        Picture picture => picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
        GraphicFrame frame => frame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
        _ => null
    };
}
