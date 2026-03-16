using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxMcp.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Services;

public class PresentationService
{
    public IReadOnlyList<SlideInfo> GetSlides(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var presentation = doc.PresentationPart!.Presentation;
        var slideIdList = presentation.SlideIdList;
        if (slideIdList is null) return [];

        var result = new List<SlideInfo>();
        int index = 0;
        foreach (SlideId slideId in slideIdList.Elements<SlideId>())
        {
            var slidePart = (SlidePart)doc.PresentationPart.GetPartById(slideId.RelationshipId!.Value!);
            var slide = slidePart.Slide;

            string? title = GetSlideTitle(slide);
            string? notes = GetSlideNotes(slidePart);
            int placeholderCount = GetPlaceholderCount(slide);

            result.Add(new SlideInfo(index, title, notes, placeholderCount));
            index++;
        }
        return result;
    }

    private static string? GetSlideTitle(Slide slide)
    {
        // Title placeholder has type "ctrTitle" or "title" or idx=0
        foreach (var shape in slide.CommonSlideData!.ShapeTree!.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
            if (ph is not null)
            {
                var phType = ph.Type?.Value;
                if (phType == PlaceholderValues.Title || phType == PlaceholderValues.CenteredTitle)
                {
                    return shape.TextBody?.InnerText;
                }
            }
        }
        // Fallback: first placeholder
        foreach (var shape in slide.CommonSlideData!.ShapeTree!.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
            if (ph is not null && (ph.Index is null || ph.Index.Value == 0))
            {
                return shape.TextBody?.InnerText;
            }
        }
        return null;
    }

    private static string? GetSlideNotes(SlidePart slidePart)
    {
        if (slidePart.NotesSlidePart is null) return null;
        var notesSlide = slidePart.NotesSlidePart.NotesSlide;
        // Notes body placeholder has type "body" or idx=1
        foreach (var shape in notesSlide.CommonSlideData!.ShapeTree!.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
            if (ph is not null && ph.Type?.Value == PlaceholderValues.Body)
            {
                var text = shape.TextBody?.InnerText;
                return string.IsNullOrEmpty(text) ? null : text;
            }
        }
        return null;
    }

    private static int GetPlaceholderCount(Slide slide)
    {
        int count = 0;
        foreach (var shape in slide.CommonSlideData!.ShapeTree!.Elements<Shape>())
        {
            if (shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape is not null)
                count++;
        }
        return count;
    }

    public IReadOnlyList<SlideLayoutInfo> GetLayouts(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var result = new List<SlideLayoutInfo>();
        int index = 0;
        foreach (var masterPart in doc.PresentationPart!.SlideMasterParts)
        {
            foreach (var layoutPart in masterPart.SlideLayoutParts)
            {
                var name = layoutPart.SlideLayout.CommonSlideData?.Name?.Value ?? $"Layout {index}";
                result.Add(new SlideLayoutInfo(index, name));
                index++;
            }
        }
        return result;
    }

    public int AddSlide(string filePath, string? layoutName)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var presentationPart = doc.PresentationPart!;
        var presentation = presentationPart.Presentation;

        // Find layout
        SlideLayoutPart? targetLayoutPart = null;
        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            foreach (var layoutPart in masterPart.SlideLayoutParts)
            {
                var name = layoutPart.SlideLayout.CommonSlideData?.Name?.Value;
                if (layoutName is null || name == layoutName)
                {
                    targetLayoutPart = layoutPart;
                    break;
                }
            }
            if (targetLayoutPart is not null) break;
        }

        // Fallback to very first layout if not found
        if (targetLayoutPart is null)
        {
            targetLayoutPart = presentationPart.SlideMasterParts.First().SlideLayoutParts.First();
        }

        // Create new slide part
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        var slide = new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()));

        slidePart.Slide = slide;
        slidePart.AddPart(targetLayoutPart);

        // Add to presentation slide list
        var slideIdList = presentation.SlideIdList ??= new SlideIdList();
        uint maxId = slideIdList.Elements<SlideId>().Any()
            ? slideIdList.Elements<SlideId>().Max(s => s.Id!.Value)
            : 255;
        var newSlideId = new SlideId
        {
            Id = maxId + 1,
            RelationshipId = presentationPart.GetIdOfPart(slidePart)
        };
        slideIdList.Append(newSlideId);

        presentation.Save();

        return slideIdList.Elements<SlideId>().ToList().IndexOf(newSlideId);
    }

    public void UpdateTextPlaceholder(string filePath, int slideIndex, int placeholderIndex, string text)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var slidePart = GetSlidePart(doc, slideIndex);
        var slide = slidePart.Slide;

        var placeholders = slide.CommonSlideData!.ShapeTree!.Elements<Shape>()
            .Where(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape is not null)
            .ToList();

        if (placeholderIndex < 0 || placeholderIndex >= placeholders.Count)
            throw new ArgumentOutOfRangeException(nameof(placeholderIndex), $"Placeholder index {placeholderIndex} is out of range. Slide has {placeholders.Count} placeholder(s).");

        var shape = placeholders[placeholderIndex];
        var textBody = shape.TextBody;
        if (textBody is null)
        {
            textBody = new TextBody(new A.BodyProperties(), new A.Paragraph());
            shape.Append(textBody);
        }

        // Replace all paragraphs with single paragraph containing the text
        foreach (var para in textBody.Elements<A.Paragraph>().ToList())
            para.Remove();

        var paragraph = new A.Paragraph(new A.Run(new A.Text(text)));
        textBody.Append(paragraph);

        slide.Save();
    }

    public void InsertImage(string filePath, int slideIndex, string imagePath, long x, long y, long width, long height)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var slidePart = GetSlidePart(doc, slideIndex);

        var imagePartType = GetImagePartType(imagePath);

        var imagePart = slidePart.AddImagePart(imagePartType);
        using (var stream = File.OpenRead(imagePath))
            imagePart.FeedData(stream);

        var imageRelId = slidePart.GetIdOfPart(imagePart);

        // Get next shape ID by scanning all shape-tree children that carry NonVisualDrawingProperties
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
        uint maxId = GetMaxShapeId(shapeTree);
        uint newId = maxId + 1;

        var picture = new Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = newId, Name = $"Image{newId}" },
                new P.NonVisualPictureDrawingProperties(
                    new A.PictureLocks { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = imageRelId },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = width, Cy = height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        shapeTree.Append(picture);
        slidePart.Slide.Save();
    }

    public string GetSlideXml(string filePath, int slideIndex)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var slidePart = GetSlidePart(doc, slideIndex);
        using var ms = new System.IO.MemoryStream();
        slidePart.Slide.Save(ms);
        return System.Text.Encoding.UTF8.GetString(ms.ToArray());
    }

    private static uint GetMaxShapeId(ShapeTree shapeTree)
    {
        uint maxId = 1;
        foreach (var child in shapeTree.ChildElements)
        {
            uint? id = child switch
            {
                Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
                P.GraphicFrame gf => gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
                P.GroupShape gs => gs.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                P.ConnectionShape cs => cs.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                _ => null
            };
            if (id.HasValue && id.Value > maxId) maxId = id.Value;
        }
        return maxId;
    }

    private static SlidePart GetSlidePart(PresentationDocument doc, int slideIndex)
    {
        var slideIdList = doc.PresentationPart!.Presentation.SlideIdList
            ?? throw new InvalidOperationException("Presentation has no slides.");
        var slideIds = slideIdList.Elements<SlideId>().ToList();
        if (slideIndex < 0 || slideIndex >= slideIds.Count)
            throw new ArgumentOutOfRangeException(nameof(slideIndex), $"Slide index {slideIndex} is out of range. Presentation has {slideIds.Count} slide(s).");
        return (SlidePart)doc.PresentationPart.GetPartById(slideIds[slideIndex].RelationshipId!.Value!);
    }

    private static PartTypeInfo GetImagePartType(string imagePath) =>
        Path.GetExtension(imagePath).ToLowerInvariant() switch
        {
            ".png" => ImagePartType.Png,
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            _ => ImagePartType.Png
        };

    public SlideContent GetSlideContent(string filePath, int slideIndex)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var presentationPart = doc.PresentationPart!;
        var slidePart = GetSlidePart(doc, slideIndex);

        // Slide dimensions from the presentation-level SlideSize element
        var slideSize = presentationPart.Presentation.SlideSize;
        long slideWidth = slideSize?.Cx?.Value ?? 9144000;
        long slideHeight = slideSize?.Cy?.Value ?? 6858000;

        var shapes = new List<ShapeContent>();
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree is not null)
        {
            foreach (var element in shapeTree.ChildElements)
            {
                var shape = ExtractShape(element);
                if (shape is not null)
                    shapes.Add(shape);
            }
        }

        return new SlideContent(slideIndex, slideWidth, slideHeight, shapes);
    }

    private static ShapeContent? ExtractShape(DocumentFormat.OpenXml.OpenXmlElement element)
    {
        return element switch
        {
            Shape s => ExtractTextShape(s),
            Picture p => ExtractPicture(p),
            P.GraphicFrame gf => ExtractGraphicFrame(gf),
            P.GroupShape gs => ExtractGroupShape(gs),
            P.ConnectionShape cs => ExtractConnectionShape(cs),
            _ => null
        };
    }

    private static ShapeContent ExtractTextShape(Shape shape)
    {
        var nvProps = shape.NonVisualShapeProperties;
        var drawingProps = nvProps?.NonVisualDrawingProperties;
        var appProps = nvProps?.ApplicationNonVisualDrawingProperties;
        var ph = appProps?.PlaceholderShape;
        var xfrm = shape.ShapeProperties?.Transform2D;

        var paragraphs = new List<string>();
        if (shape.TextBody is not null)
        {
            foreach (var para in shape.TextBody.Elements<A.Paragraph>())
                paragraphs.Add(para.InnerText);
        }

        return new ShapeContent(
            ShapeId: drawingProps?.Id?.Value,
            Name: drawingProps?.Name?.Value ?? "",
            ShapeType: "Text",
            X: xfrm?.Offset?.X?.Value,
            Y: xfrm?.Offset?.Y?.Value,
            Width: xfrm?.Extents?.Cx?.Value,
            Height: xfrm?.Extents?.Cy?.Value,
            IsPlaceholder: ph is not null,
            PlaceholderType: ph?.Type?.Value.ToString(),
            PlaceholderIndex: ph?.Index?.Value,
            Text: paragraphs.Count > 0 ? string.Join("\n", paragraphs) : null,
            Paragraphs: paragraphs.Count > 0 ? paragraphs : null,
            TableRows: null);
    }

    private static ShapeContent ExtractPicture(Picture picture)
    {
        var nvProps = picture.NonVisualPictureProperties;
        var drawingProps = nvProps?.NonVisualDrawingProperties;
        var xfrm = picture.ShapeProperties?.Transform2D;

        return new ShapeContent(
            ShapeId: drawingProps?.Id?.Value,
            Name: drawingProps?.Name?.Value ?? "",
            ShapeType: "Picture",
            X: xfrm?.Offset?.X?.Value,
            Y: xfrm?.Offset?.Y?.Value,
            Width: xfrm?.Extents?.Cx?.Value,
            Height: xfrm?.Extents?.Cy?.Value,
            IsPlaceholder: false,
            PlaceholderType: null,
            PlaceholderIndex: null,
            Text: null,
            Paragraphs: null,
            TableRows: null);
    }

    private static ShapeContent ExtractGraphicFrame(P.GraphicFrame frame)
    {
        var nvProps = frame.NonVisualGraphicFrameProperties;
        var drawingProps = nvProps?.NonVisualDrawingProperties;
        var xfrm = frame.Transform;

        // Try to extract table content
        IReadOnlyList<IReadOnlyList<string>>? tableRows = null;
        var graphic = frame.Graphic;
        var graphicData = graphic?.GraphicData;
        if (graphicData is not null)
        {
            var table = graphicData.GetFirstChild<A.Table>();
            if (table is not null)
                tableRows = ExtractTableRows(table);
        }

        return new ShapeContent(
            ShapeId: drawingProps?.Id?.Value,
            Name: drawingProps?.Name?.Value ?? "",
            ShapeType: tableRows is not null ? "Table" : "GraphicFrame",
            X: xfrm?.Offset?.X?.Value,
            Y: xfrm?.Offset?.Y?.Value,
            Width: xfrm?.Extents?.Cx?.Value,
            Height: xfrm?.Extents?.Cy?.Value,
            IsPlaceholder: false,
            PlaceholderType: null,
            PlaceholderIndex: null,
            Text: null,
            Paragraphs: null,
            TableRows: tableRows);
    }

    private static IReadOnlyList<IReadOnlyList<string>> ExtractTableRows(A.Table table)
    {
        var rows = new List<IReadOnlyList<string>>();
        foreach (var row in table.Elements<A.TableRow>())
        {
            var cells = new List<string>();
            foreach (var cell in row.Elements<A.TableCell>())
                cells.Add(cell.InnerText);
            rows.Add(cells);
        }
        return rows;
    }

    private static ShapeContent ExtractGroupShape(P.GroupShape group)
    {
        var nvProps = group.NonVisualGroupShapeProperties;
        var drawingProps = nvProps?.NonVisualDrawingProperties;
        var xfrm = group.GroupShapeProperties?.TransformGroup;

        return new ShapeContent(
            ShapeId: drawingProps?.Id?.Value,
            Name: drawingProps?.Name?.Value ?? "",
            ShapeType: "Group",
            X: xfrm?.Offset?.X?.Value,
            Y: xfrm?.Offset?.Y?.Value,
            Width: xfrm?.Extents?.Cx?.Value,
            Height: xfrm?.Extents?.Cy?.Value,
            IsPlaceholder: false,
            PlaceholderType: null,
            PlaceholderIndex: null,
            Text: null,
            Paragraphs: null,
            TableRows: null);
    }

    private static ShapeContent ExtractConnectionShape(P.ConnectionShape connector)
    {
        var nvProps = connector.NonVisualConnectionShapeProperties;
        var drawingProps = nvProps?.NonVisualDrawingProperties;
        var xfrm = connector.ShapeProperties?.Transform2D;

        return new ShapeContent(
            ShapeId: drawingProps?.Id?.Value,
            Name: drawingProps?.Name?.Value ?? "",
            ShapeType: "Connector",
            X: xfrm?.Offset?.X?.Value,
            Y: xfrm?.Offset?.Y?.Value,
            Width: xfrm?.Extents?.Cx?.Value,
            Height: xfrm?.Extents?.Cy?.Value,
            IsPlaceholder: false,
            PlaceholderType: null,
            PlaceholderIndex: null,
            Text: null,
            Paragraphs: null,
            TableRows: null);
    }
}
