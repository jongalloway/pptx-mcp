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

    public MarkdownExportResult ExportMarkdown(string filePath, string? outputPath = null)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var presentationPart = doc.PresentationPart!;
        var slideIdList = presentationPart.Presentation.SlideIdList;
        var slideIds = slideIdList?.Elements<SlideId>().ToList() ?? [];

        var resolvedOutputPath = Path.GetFullPath(outputPath ?? Path.ChangeExtension(filePath, ".md")!);
        var outputDirectory = Path.GetDirectoryName(resolvedOutputPath) ?? Directory.GetCurrentDirectory();
        Directory.CreateDirectory(outputDirectory);

        var builder = new System.Text.StringBuilder();
        builder.AppendLine($"# {GetPresentationTitle(presentationPart, slideIds, filePath)}");
        if (slideIds.Count > 0)
            builder.AppendLine();

        var imageCount = 0;
        for (var slideIndex = 0; slideIndex < slideIds.Count; slideIndex++)
        {
            var slidePartForExport = GetSlidePart(doc, slideIndex);
            AppendSlideMarkdown(builder, slidePartForExport, slideIndex, resolvedOutputPath, ref imageCount);
        }

        var markdown = builder.ToString().TrimEnd() + Environment.NewLine;
        File.WriteAllText(resolvedOutputPath, markdown);

        return new MarkdownExportResult(resolvedOutputPath, markdown, slideIds.Count, imageCount);
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

    private static string GetPresentationTitle(PresentationPart presentationPart, IReadOnlyList<SlideId> slideIds, string filePath)
    {
        if (slideIds.Count > 0)
        {
            var firstSlidePart = (SlidePart)presentationPart.GetPartById(slideIds[0].RelationshipId!.Value!);
            var title = NormalizeMarkdownText(GetSlideTitle(firstSlidePart.Slide));
            if (!string.IsNullOrWhiteSpace(title))
                return title;
        }

        return NormalizeMarkdownText(Path.GetFileNameWithoutExtension(filePath)) ?? "Presentation";
    }

    private static void AppendSlideMarkdown(System.Text.StringBuilder builder, SlidePart slidePart, int slideIndex, string outputPath, ref int imageCount)
    {
        var slideTitle = NormalizeMarkdownText(GetSlideTitle(slidePart.Slide)) ?? $"Untitled Slide {slideIndex + 1}";
        builder.AppendLine($"## Slide {slideIndex + 1}: {slideTitle}");
        builder.AppendLine();

        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree is null)
            return;

        var slideImageCount = 0;
        foreach (var element in shapeTree.ChildElements)
        {
            switch (element)
            {
                case Shape shape when AppendTextShapeMarkdown(builder, shape):
                    builder.AppendLine();
                    break;
                case Picture picture:
                {
                    slideImageCount++;
                    var imageMarkdown = ExportPictureMarkdown(slidePart, picture, slideIndex, slideImageCount, outputPath);
                    if (imageMarkdown is not null)
                    {
                        imageCount++;
                        builder.AppendLine(imageMarkdown);
                        builder.AppendLine();
                    }
                    break;
                }
                case P.GraphicFrame frame:
                {
                    var table = frame.Graphic?.GraphicData?.GetFirstChild<A.Table>();
                    if (table is not null && AppendTableMarkdown(builder, table))
                        builder.AppendLine();
                    break;
                }
            }
        }
    }

    private static bool AppendTextShapeMarkdown(System.Text.StringBuilder builder, Shape shape)
    {
        if (shape.TextBody is null)
            return false;

        var paragraphs = shape.TextBody.Elements<A.Paragraph>()
            .Select(paragraph => new
            {
                Paragraph = paragraph,
                Text = NormalizeMarkdownText(paragraph.InnerText)
            })
            .Where(item => !string.IsNullOrWhiteSpace(item.Text))
            .ToList();

        if (paragraphs.Count == 0)
            return false;

        var placeholderType = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value;
        if (placeholderType == PlaceholderValues.Title || placeholderType == PlaceholderValues.CenteredTitle)
            return false;

        if (placeholderType == PlaceholderValues.SubTitle)
        {
            foreach (var paragraph in paragraphs)
                builder.AppendLine($"### {paragraph.Text}");
            return true;
        }

        var treatAsList = paragraphs.Any(item => IsExplicitListParagraph(item.Paragraph))
            || (placeholderType == PlaceholderValues.Body && paragraphs.Count > 1);

        foreach (var paragraph in paragraphs)
        {
            if (ShouldRenderAsListItem(paragraph.Paragraph, treatAsList))
            {
                var level = GetParagraphLevel(paragraph.Paragraph);
                var marker = IsNumberedParagraph(paragraph.Paragraph) ? "1." : "-";
                builder.AppendLine($"{new string(' ', level * 2)}{marker} {paragraph.Text}");
                continue;
            }

            builder.AppendLine(paragraph.Text!);
        }

        return true;
    }

    private static bool AppendTableMarkdown(System.Text.StringBuilder builder, A.Table table)
    {
        var rows = ExtractTableRows(table)
            .Select(row => row.Select(EscapeMarkdownTableCell).ToList())
            .Where(row => row.Count > 0)
            .ToList();

        if (rows.Count == 0)
            return false;

        var columnCount = rows.Max(row => row.Count);
        var normalizedRows = rows
            .Select(row => row.Concat(Enumerable.Repeat(string.Empty, columnCount - row.Count)).ToList())
            .ToList();

        builder.AppendLine($"| {string.Join(" | ", normalizedRows[0])} |");
        builder.AppendLine($"| {string.Join(" | ", Enumerable.Repeat("---", columnCount))} |");
        foreach (var row in normalizedRows.Skip(1))
            builder.AppendLine($"| {string.Join(" | ", row)} |");

        return true;
    }

    private static string? ExportPictureMarkdown(SlidePart slidePart, Picture picture, int slideIndex, int imageIndex, string outputPath)
    {
        var relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId))
            return null;

        if (slidePart.GetPartById(relationshipId) is not ImagePart imagePart)
            return null;

        var outputDirectory = Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory();
        var imageDirectoryName = $"{Path.GetFileNameWithoutExtension(outputPath)}_images";
        var imageDirectoryPath = Path.Join(outputDirectory, imageDirectoryName);
        Directory.CreateDirectory(imageDirectoryPath);

        var altText = NormalizeMarkdownText(picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value)
            ?? $"Slide {slideIndex + 1} image {imageIndex}";
        var extension = GetImageExtension(imagePart);
        var imageFileName = $"slide-{slideIndex + 1}-image-{imageIndex}-{SanitizeFileName(altText)}{extension}";
        var imagePath = GetUniquePath(imageDirectoryPath, imageFileName);

        using (var imageStream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
        using (var outputStream = File.Create(imagePath))
        {
            imageStream.CopyTo(outputStream);
        }

        var relativePath = Path.GetRelativePath(outputDirectory, imagePath).Replace('\\', '/');
        return $"![{altText}]({relativePath})";
    }

    private static bool ShouldRenderAsListItem(A.Paragraph paragraph, bool treatAsList)
    {
        var properties = paragraph.ParagraphProperties;
        if (properties?.GetFirstChild<A.NoBullet>() is not null)
            return false;

        return IsExplicitListParagraph(paragraph) || treatAsList;
    }

    private static bool IsExplicitListParagraph(A.Paragraph paragraph)
    {
        var properties = paragraph.ParagraphProperties;
        if (properties is null)
            return false;

        return properties.GetFirstChild<A.CharacterBullet>() is not null
            || properties.GetFirstChild<A.AutoNumberedBullet>() is not null
            || properties.Level is not null;
    }

    private static bool IsNumberedParagraph(A.Paragraph paragraph) =>
        paragraph.ParagraphProperties?.GetFirstChild<A.AutoNumberedBullet>() is not null;

    private static int GetParagraphLevel(A.Paragraph paragraph) =>
        paragraph.ParagraphProperties?.Level?.Value is { } level ? level : 0;

    private static string EscapeMarkdownTableCell(string value) =>
        (NormalizeMarkdownText(value) ?? string.Empty).Replace("|", "\\|");

    private static string? NormalizeMarkdownText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        return value
            .Replace("\r", string.Empty)
            .Replace("\n", " ")
            .Trim();
    }

    private static string GetImageExtension(ImagePart imagePart)
    {
        var uriExtension = Path.GetExtension(imagePart.Uri.ToString());
        if (!string.IsNullOrWhiteSpace(uriExtension))
            return uriExtension;

        return imagePart.ContentType switch
        {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            _ => ".bin"
        };
    }

    private static string SanitizeFileName(string value)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        var builder = new System.Text.StringBuilder(value.Length);
        foreach (var character in value)
            builder.Append(invalidCharacters.Contains(character) ? '_' : character);

        return builder.ToString().Trim() switch
        {
            "" => "image",
            var sanitized => sanitized
        };
    }

    private static string GetUniquePath(string directory, string fileName)
    {
        var candidate = Path.Join(directory, fileName);
        if (!File.Exists(candidate))
            return candidate;

        var baseName = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);
        var counter = 2;
        while (true)
        {
            candidate = Path.Join(directory, $"{baseName}-{counter}{extension}");
            if (!File.Exists(candidate))
                return candidate;
            counter++;
        }
    }
}
