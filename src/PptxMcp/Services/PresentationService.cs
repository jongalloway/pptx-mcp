using DocumentFormat.OpenXml;
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
                var paraTexts = shape.TextBody?
                    .Elements<A.Paragraph>()
                    .Select(p => p.InnerText)
                    .ToList();
                if (paraTexts is not { Count: > 0 }) return null;
                var text = string.Join("\n", paraTexts);
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

        ReplaceShapeTextPreservingFormatting(placeholders[placeholderIndex], text);
        slide.Save();
    }

    public SlideDataUpdateResult UpdateSlideData(string filePath, int slideNumber, string? shapeName, int? placeholderIndex, string newText)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var slideIds = GetSlideIds(doc);

        var result = UpdateSlideData(doc, slideIds, slideNumber, shapeName, placeholderIndex, newText, out var modifiedSlidePart);
        if (result.Success)
            modifiedSlidePart?.Slide?.Save();

        return result;
    }

    public BatchUpdateResult BatchUpdate(string filePath, IReadOnlyList<BatchUpdateMutation> mutations)
    {
        if (mutations is null || mutations.Count == 0)
            return new BatchUpdateResult(0, 0, 0, []);

        using var doc = PresentationDocument.Open(filePath, true);
        var slideIds = GetSlideIds(doc);
        var results = new List<BatchUpdateMutationResult>(mutations.Count);
        var modifiedSlideParts = new HashSet<SlidePart>();

        foreach (var mutation in mutations)
        {
            var mutationResult = UpdateSlideData(doc, slideIds, mutation.SlideNumber, mutation.ShapeName, null, mutation.NewValue, out var modifiedSlidePart);
            if (mutationResult.Success && modifiedSlidePart is not null)
                modifiedSlideParts.Add(modifiedSlidePart);

            results.Add(new BatchUpdateMutationResult(
                SlideNumber: mutation.SlideNumber,
                ShapeName: mutation.ShapeName,
                Success: mutationResult.Success,
                Error: mutationResult.Success ? null : mutationResult.Message,
                MatchedBy: mutationResult.MatchedBy));
        }

        foreach (var modifiedSlidePart in modifiedSlideParts)
            modifiedSlidePart.Slide?.Save();

        var successCount = results.Count(result => result.Success);
        return new BatchUpdateResult(
            TotalMutations: results.Count,
            SuccessCount: successCount,
            FailureCount: results.Count - successCount,
            Results: results);
    }

    public void WriteNotes(string filePath, int slideIndex, string notes, bool append = false)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var presentationPart = doc.PresentationPart!;
        var slidePart = GetSlidePart(doc, slideIndex);

        var notesMasterPart = EnsureNotesMasterPart(presentationPart);
        var lines = notes.Split('\n');

        if (slidePart.NotesSlidePart is null)
        {
            CreateNotesSlidePart(slidePart, notesMasterPart, lines);
        }
        else
        {
            if (slidePart.NotesSlidePart.NotesMasterPart is null)
                slidePart.NotesSlidePart.AddPart(notesMasterPart);
            if (append)
            {
                var existing = GetSlideNotes(slidePart) ?? string.Empty;
                var combined = string.IsNullOrEmpty(existing) ? notes : existing + "\n" + notes;
                lines = combined.Split('\n');
            }
            UpdateNotesSlideContent(slidePart.NotesSlidePart, lines);
        }
    }

    private static NotesMasterPart EnsureNotesMasterPart(PresentationPart presentationPart)
    {
        if (presentationPart.NotesMasterPart is not null)
            return presentationPart.NotesMasterPart;

        var notesMasterPart = presentationPart.AddNewPart<NotesMasterPart>();
        notesMasterPart.NotesMaster = new NotesMaster(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            });
        notesMasterPart.NotesMaster.Save();
        return notesMasterPart;
    }

    private static void CreateNotesSlidePart(SlidePart slidePart, NotesMasterPart notesMasterPart, string[] paragraphs)
    {
        var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
        notesSlidePart.NotesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()),
                    BuildNotesBodyShape(paragraphs))),
            new ColorMapOverride(new A.MasterColorMapping()));
        notesSlidePart.AddPart(slidePart);
        notesSlidePart.AddPart(notesMasterPart);
        notesSlidePart.NotesSlide.Save();
    }

    private static void UpdateNotesSlideContent(NotesSlidePart notesSlidePart, string[] paragraphs)
    {
        var notesSlide = notesSlidePart.NotesSlide;
        var shapeTree = notesSlide.CommonSlideData!.ShapeTree!;
        foreach (var shape in shapeTree.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
            if (ph is not null && ph.Type?.Value == PlaceholderValues.Body)
            {
                if (shape.TextBody is null)
                    shape.Append(new TextBody(new A.BodyProperties(), new A.ListStyle()));
                var textBody = shape.TextBody!;
                foreach (var para in textBody.Elements<A.Paragraph>().ToList())
                    para.Remove();
                foreach (var line in paragraphs)
                    textBody.Append(new A.Paragraph(
                        new A.Run(new A.Text(line)),
                        new A.EndParagraphRunProperties()));
                notesSlide.Save();
                return;
            }
        }
        // Body placeholder not found — add one to make the write reliable
        shapeTree.Append(BuildNotesBodyShape(paragraphs));
        notesSlide.Save();
    }

    private static Shape BuildNotesBodyShape(string[] paragraphs)
    {
        var textBody = new TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var line in paragraphs)
            textBody.Append(new A.Paragraph(
                new A.Run(new A.Text(line)),
                new A.EndParagraphRunProperties()));
        return new Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2U, Name = "Notes Placeholder 2" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1U })),
            new ShapeProperties(),
            textBody);
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
        var slideIds = GetSlideIds(doc);
        return GetSlidePart(doc, slideIds, slideIndex);
    }

    private static SlidePart GetSlidePart(PresentationDocument doc, IReadOnlyList<SlideId> slideIds, int slideIndex)
    {
        if (slideIds.Count == 0)
            throw new InvalidOperationException("Presentation has no slides.");
        if (slideIndex < 0 || slideIndex >= slideIds.Count)
            throw new ArgumentOutOfRangeException(nameof(slideIndex), $"Slide index {slideIndex} is out of range. Presentation has {slideIds.Count} slide(s).");
        return (SlidePart)doc.PresentationPart!.GetPartById(slideIds[slideIndex].RelationshipId!.Value!);
    }

    private static IReadOnlyList<SlideId> GetSlideIds(PresentationDocument doc) =>
        doc.PresentationPart!.Presentation.SlideIdList?.Elements<SlideId>().ToList() ?? [];

    private static SlideDataUpdateResult UpdateSlideData(
        PresentationDocument doc,
        IReadOnlyList<SlideId> slideIds,
        int slideNumber,
        string? shapeName,
        int? placeholderIndex,
        string newText,
        out SlidePart? modifiedSlidePart)
    {
        modifiedSlidePart = null;

        if (slideNumber <= 0)
            return CreateSlideDataUpdateFailure(slideNumber, shapeName, placeholderIndex, newText, "slideNumber must be 1 or greater.");

        if (placeholderIndex is < 0)
            return CreateSlideDataUpdateFailure(slideNumber, shapeName, placeholderIndex, newText, "placeholderIndex must be zero or greater.");

        if (string.IsNullOrWhiteSpace(shapeName) && placeholderIndex is null)
            return CreateSlideDataUpdateFailure(slideNumber, shapeName, placeholderIndex, newText, "Provide either shapeName or placeholderIndex.");

        if (slideIds.Count == 0)
            return CreateSlideDataUpdateFailure(slideNumber, shapeName, placeholderIndex, newText, "Presentation has no slides.");

        if (slideNumber > slideIds.Count)
            return CreateSlideDataUpdateFailure(slideNumber, shapeName, placeholderIndex, newText, $"slideNumber {slideNumber} is out of range. Presentation has {slideIds.Count} slide(s).");

        var slidePart = GetSlidePart(doc, slideIds, slideNumber - 1);
        if (slidePart.Slide is null)
            return CreateSlideDataUpdateFailure(slideNumber, shapeName, placeholderIndex, newText, $"Slide {slideNumber} could not be loaded.");

        var textShapeTargets = GetTextShapeTargets(slidePart.Slide).ToList();
        if (textShapeTargets.Count == 0)
            return CreateSlideDataUpdateFailure(slideNumber, shapeName, placeholderIndex, newText, $"Slide {slideNumber} does not contain any text-capable shapes.");

        var target = ResolveTextShapeTarget(textShapeTargets, shapeName, placeholderIndex, out var matchedBy, out var failureMessage);
        if (target is null)
            return CreateSlideDataUpdateFailure(slideNumber, shapeName, placeholderIndex, newText, failureMessage ?? "Unable to resolve a target shape.");

        var previousText = GetShapeText(target.Shape);
        ReplaceShapeTextPreservingFormatting(target.Shape, newText);
        modifiedSlidePart = slidePart;

        return new SlideDataUpdateResult(
            Success: true,
            SlideNumber: slideNumber,
            RequestedShapeName: shapeName,
            RequestedPlaceholderIndex: placeholderIndex,
            MatchedBy: matchedBy,
            ResolvedShapeName: target.Name,
            ResolvedShapeIndex: target.Index,
            ResolvedShapeId: target.ShapeId,
            PlaceholderType: target.PlaceholderType,
            LayoutPlaceholderIndex: target.LayoutPlaceholderIndex,
            PreviousText: previousText,
            NewText: newText,
            Message: $"Updated shape '{target.Name}' on slide {slideNumber}.");
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

    private static SlideDataUpdateResult CreateSlideDataUpdateFailure(int slideNumber, string? shapeName, int? placeholderIndex, string newText, string message) =>
        new(
            Success: false,
            SlideNumber: slideNumber,
            RequestedShapeName: shapeName,
            RequestedPlaceholderIndex: placeholderIndex,
            MatchedBy: null,
            ResolvedShapeName: null,
            ResolvedShapeIndex: null,
            ResolvedShapeId: null,
            PlaceholderType: null,
            LayoutPlaceholderIndex: null,
            PreviousText: null,
            NewText: newText,
            Message: message);

    private static IReadOnlyList<TextShapeTarget> GetTextShapeTargets(Slide slide)
    {
        var shapeTree = slide.CommonSlideData?.ShapeTree;
        if (shapeTree is null)
            return [];

        return shapeTree.Elements<Shape>()
            .Select((shape, index) =>
            {
                var drawingProperties = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
                var placeholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
                var name = drawingProperties?.Name?.Value;

                return new TextShapeTarget(
                    Shape: shape,
                    Index: index,
                    Name: string.IsNullOrWhiteSpace(name) ? $"Shape {index}" : name,
                    ShapeId: drawingProperties?.Id?.Value,
                    PlaceholderType: placeholderShape?.Type?.Value.ToString(),
                    LayoutPlaceholderIndex: placeholderShape?.Index?.Value);
            })
            .ToList();
    }

    private static TextShapeTarget? ResolveTextShapeTarget(
        IReadOnlyList<TextShapeTarget> textShapeTargets,
        string? shapeName,
        int? placeholderIndex,
        out string? matchedBy,
        out string? failureMessage)
    {
        matchedBy = null;
        failureMessage = null;

        if (!string.IsNullOrWhiteSpace(shapeName))
        {
            var trimmedShapeName = shapeName.Trim();
            var matches = textShapeTargets
                .Where(target => string.Equals(target.Name, trimmedShapeName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (matches.Count == 1)
            {
                matchedBy = "shapeName";
                return matches[0];
            }

            if (matches.Count > 1)
            {
                failureMessage = $"Multiple text-capable shapes named '{trimmedShapeName}' were found on slide. Use placeholderIndex to disambiguate.";
                return null;
            }
        }

        if (placeholderIndex is null)
        {
            failureMessage = $"No text-capable shape named '{shapeName}' was found. Available shapes: {DescribeTextShapeTargets(textShapeTargets)}";
            return null;
        }

        if (placeholderIndex.Value >= textShapeTargets.Count)
        {
            failureMessage = $"placeholderIndex {placeholderIndex.Value} is out of range. Slide has {textShapeTargets.Count} text-capable shape(s): {DescribeTextShapeTargets(textShapeTargets)}";
            return null;
        }

        matchedBy = string.IsNullOrWhiteSpace(shapeName)
            ? "placeholderIndex"
            : "placeholderIndexFallback";
        return textShapeTargets[placeholderIndex.Value];
    }

    private static string DescribeTextShapeTargets(IEnumerable<TextShapeTarget> textShapeTargets) =>
        string.Join(", ",
            textShapeTargets.Select(target => $"{target.Index}:{target.Name}"));

    private static string? GetShapeText(Shape shape)
    {
        var paragraphs = shape.TextBody?
            .Elements<A.Paragraph>()
            .Select(paragraph => paragraph.InnerText)
            .ToList();

        return paragraphs is { Count: > 0 }
            ? string.Join("\n", paragraphs)
            : null;
    }

    private static void ReplaceShapeTextPreservingFormatting(Shape shape, string text)
    {
        var existingTextBody = shape.TextBody ?? new TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties()));
        var paragraphTemplates = existingTextBody.Elements<A.Paragraph>().ToArray();
        if (paragraphTemplates.Length == 0)
            paragraphTemplates = [new A.Paragraph(new A.EndParagraphRunProperties())];

        var replacementTextBody = new TextBody(
            existingTextBody.BodyProperties is null
                ? new A.BodyProperties()
                : (A.BodyProperties)existingTextBody.BodyProperties.CloneNode(true),
            existingTextBody.ListStyle is null
                ? new A.ListStyle()
                : (A.ListStyle)existingTextBody.ListStyle.CloneNode(true));

        var replacementParagraphs = GetReplacementParagraphs(text);
        for (var index = 0; index < replacementParagraphs.Count; index++)
        {
            var template = paragraphTemplates[Math.Min(index, paragraphTemplates.Length - 1)];
            replacementTextBody.Append(CreateParagraphFromTemplate(template, replacementParagraphs[index]));
        }

        var existingTextBodyElement = shape.TextBody;
        if (existingTextBodyElement is not null)
        {
            shape.ReplaceChild(replacementTextBody, existingTextBodyElement);
            return;
        }

        DocumentFormat.OpenXml.OpenXmlElement? insertAfter = null;
        var shapeProperties = shape.GetFirstChild<P.ShapeProperties>();
        var shapeStyle = shape.GetFirstChild<P.ShapeStyle>();

        if (shapeStyle is not null)
            insertAfter = shapeStyle;
        else if (shapeProperties is not null)
            insertAfter = shapeProperties;

        if (insertAfter is not null)
        {
            shape.InsertAfter(replacementTextBody, insertAfter);
            return;
        }

        var extensionList = shape.GetFirstChild<P.ExtensionList>();
        if (extensionList is not null)
            shape.InsertBefore(replacementTextBody, extensionList);
        else
            shape.Append(replacementTextBody);
    }

    private static IReadOnlyList<string> GetReplacementParagraphs(string text)
    {
        var normalizedText = text
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');

        return normalizedText.Split('\n', StringSplitOptions.None);
    }

    private static A.Paragraph CreateParagraphFromTemplate(A.Paragraph template, string text)
    {
        var paragraph = new A.Paragraph();
        if (template.ParagraphProperties is not null)
            paragraph.Append((A.ParagraphProperties)template.ParagraphProperties.CloneNode(true));

        var runTemplate = template.Elements<A.Run>().FirstOrDefault()?.RunProperties;
        var runProperties = runTemplate is null
            ? new A.RunProperties()
            : (A.RunProperties)runTemplate.CloneNode(true);

        var textElement = new A.Text(text);
        textElement.SetAttribute(new OpenXmlAttribute("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve"));

        paragraph.Append(new A.Run(runProperties, textElement));

        var templateEndParagraphRunProperties = template.Elements<A.EndParagraphRunProperties>().FirstOrDefault();
        var endParagraphRunProperties = templateEndParagraphRunProperties is null
            ? new A.EndParagraphRunProperties()
            : (A.EndParagraphRunProperties)templateEndParagraphRunProperties.CloneNode(true);
        paragraph.Append(endParagraphRunProperties);

        return paragraph;
    }

    private sealed record TextShapeTarget(
        Shape Shape,
        int Index,
        string Name,
        uint? ShapeId,
        string? PlaceholderType,
        uint? LayoutPlaceholderIndex);

    public IReadOnlyList<SlideTalkingPoints> ExtractTalkingPoints(string filePath, int topN = 5)
    {
        if (topN <= 0)
            throw new ArgumentOutOfRangeException(nameof(topN), "topN must be greater than zero.");

        using var doc = PresentationDocument.Open(filePath, false);
        var slideIdList = doc.PresentationPart!.Presentation.SlideIdList;
        if (slideIdList is null)
            return [];

        var talkingPoints = new List<SlideTalkingPoints>();
        var slideIds = slideIdList.Elements<SlideId>().ToList();
        for (int slideIndex = 0; slideIndex < slideIds.Count; slideIndex++)
        {
            var slidePart = GetSlidePart(doc, slideIndex);
            try
            {
                var slideContent = GetSlideContent(doc.PresentationPart!, slidePart, slideIndex);
                var title = GetSlideTitle(slidePart.Slide);
                var points = RankTalkingPoints(slideContent, topN);
                talkingPoints.Add(new SlideTalkingPoints(slideIndex, title, points));
            }
            catch
            {
                talkingPoints.Add(new SlideTalkingPoints(slideIndex, null, []));
            }
        }

        return talkingPoints;
    }

    public SlideContent GetSlideContent(string filePath, int slideIndex)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var slidePart = GetSlidePart(doc, slideIndex);
        return GetSlideContent(doc.PresentationPart!, slidePart, slideIndex);
    }

    public IReadOnlyList<SlideContent> GetAllSlideContents(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var slideIdList = doc.PresentationPart!.Presentation.SlideIdList;
        if (slideIdList is null)
            return [];

        var slideIds = slideIdList.Elements<SlideId>().ToList();
        var result = new List<SlideContent>(slideIds.Count);
        for (int i = 0; i < slideIds.Count; i++)
        {
            var slidePart = GetSlidePart(doc, i);
            result.Add(GetSlideContent(doc.PresentationPart!, slidePart, i));
        }
        return result;
    }

    public MarkdownExportResult ExportMarkdown(string filePath, string? outputPath = null)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var presentationPart = doc.PresentationPart!;
        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<SlideId>().ToList() ?? [];

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
            var slidePart = GetSlidePart(doc, slideIndex);
            AppendSlideMarkdown(builder, slidePart, slideIndex, resolvedOutputPath, ref imageCount);
        }

        var markdown = builder.ToString().TrimEnd() + Environment.NewLine;
        File.WriteAllText(resolvedOutputPath, markdown);

        return new MarkdownExportResult(resolvedOutputPath, markdown, slideIds.Count, imageCount);
    }

    private static SlideContent GetSlideContent(PresentationPart presentationPart, SlidePart slidePart, int slideIndex)
    {
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

    private static IReadOnlyList<string> RankTalkingPoints(SlideContent slideContent, int topN)
    {
        var candidates = new List<TalkingPointCandidate>();
        int order = 0;

        foreach (var shape in slideContent.Shapes)
        {
            if (!string.Equals(shape.ShapeType, "Text", StringComparison.OrdinalIgnoreCase) || shape.Paragraphs is null)
                continue;

            if (ShouldIgnoreShape(shape))
                continue;

            for (int paragraphIndex = 0; paragraphIndex < shape.Paragraphs.Count; paragraphIndex++)
            {
                var text = NormalizeText(shape.Paragraphs[paragraphIndex]);
                if (ShouldIgnoreParagraph(text))
                    continue;

                var score = ScoreTalkingPoint(shape, text);
                if (score <= 0)
                    continue;

                candidates.Add(new TalkingPointCandidate(
                    text,
                    NormalizeKey(text),
                    score,
                    shape.Y ?? long.MaxValue,
                    order++,
                    IsTitlePlaceholder(shape.PlaceholderType)));
            }
        }

        var selectedCandidates = candidates
            .GroupBy(candidate => candidate.Key, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.OrderByDescending(candidate => candidate.Score).ThenBy(candidate => candidate.Order).First())
            .OrderByDescending(candidate => candidate.Score)
            .ThenBy(candidate => candidate.Y)
            .ThenBy(candidate => candidate.Order)
            .Take(topN)
            .ToList();

        var hasVisualOnlyContent = slideContent.Shapes.Any(shape =>
            !string.Equals(shape.ShapeType, "Text", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(shape.ShapeType, "Connector", StringComparison.OrdinalIgnoreCase));

        if (hasVisualOnlyContent && selectedCandidates.Count > 0 && selectedCandidates.All(candidate => candidate.IsTitleCandidate))
            return [];

        return selectedCandidates
            .OrderBy(candidate => candidate.Y)
            .ThenBy(candidate => candidate.Order)
            .Select(candidate => candidate.Text)
            .ToList();
    }

    private static bool ShouldIgnoreShape(ShapeContent shape)
    {
        if (ContainsNoiseMarker(shape.Name))
            return true;

        return shape.PlaceholderType switch
        {
            nameof(PlaceholderValues.DateAndTime) => true,
            nameof(PlaceholderValues.Footer) => true,
            nameof(PlaceholderValues.Header) => true,
            nameof(PlaceholderValues.SlideNumber) => true,
            _ => false
        };
    }

    private static bool ShouldIgnoreParagraph(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return true;

        if (!text.Any(char.IsLetterOrDigit))
            return true;

        if (ContainsNoiseMarker(text))
            return true;

        return text.StartsWith("Click to add ", StringComparison.OrdinalIgnoreCase);
    }

    private static int ScoreTalkingPoint(ShapeContent shape, string text)
    {
        int score = shape.PlaceholderType switch
        {
            nameof(PlaceholderValues.Body) => 90,
            nameof(PlaceholderValues.Object) => 80,
            nameof(PlaceholderValues.SubTitle) => 60,
            nameof(PlaceholderValues.Title) => 45,
            nameof(PlaceholderValues.CenteredTitle) => 45,
            _ when shape.IsPlaceholder => 55,
            _ => 65
        };

        if ((shape.Paragraphs?.Count ?? 0) > 1)
            score += 20;

        if (LooksLikeBullet(text))
            score += 25;

        if (text.Length >= 15 && text.Length <= 160)
            score += 10;
        else if (text.Length > 200)
            score -= 10;

        if (IsLikelyHeading(text))
            score -= 15;

        return score;
    }

    private static bool LooksLikeBullet(string text)
    {
        var trimmed = text.TrimStart();
        if (trimmed.Length == 0)
            return false;

        return trimmed[0] switch
        {
            '-' => true,
            '*' => true,
            '•' => true,
            '◦' => true,
            '–' => true,
            _ => trimmed.Length > 2 && char.IsDigit(trimmed[0]) && (trimmed[1] == '.' || trimmed[1] == ')')
        };
    }

    private static bool IsLikelyHeading(string text)
    {
        if (text.Contains(':') || text.Contains('.') || text.Contains(';'))
            return false;

        return text.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length <= 5;
    }

    private static string NormalizeText(string text) =>
        string.Join(' ', text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static string NormalizeKey(string text) =>
        NormalizeText(text).TrimEnd('.', ';', ':', '!', '?').ToUpperInvariant();

    private static bool ContainsNoiseMarker(string text) =>
        text.Contains("presenter notes", StringComparison.OrdinalIgnoreCase)
        || text.Contains("speaker notes", StringComparison.OrdinalIgnoreCase)
        || text.Contains("notes placeholder", StringComparison.OrdinalIgnoreCase);

    private static bool IsTitlePlaceholder(string? placeholderType) =>
        placeholderType is nameof(PlaceholderValues.Title) or nameof(PlaceholderValues.CenteredTitle);

    private sealed record TalkingPointCandidate(string Text, string Key, int Score, long Y, int Order, bool IsTitleCandidate);

    public void MoveSlide(string filePath, int slideNumber, int targetPosition)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var presentationPart = doc.PresentationPart!;
        var slideIdList = presentationPart.Presentation.SlideIdList
            ?? throw new InvalidOperationException("Presentation has no slides.");

        var slideIds = slideIdList.Elements<SlideId>().ToList();
        int count = slideIds.Count;

        if (slideNumber < 1 || slideNumber > count)
            throw new ArgumentOutOfRangeException(nameof(slideNumber), $"slideNumber {slideNumber} is out of range. Presentation has {count} slide(s).");
        if (targetPosition < 1 || targetPosition > count)
            throw new ArgumentOutOfRangeException(nameof(targetPosition), $"targetPosition {targetPosition} is out of range. Presentation has {count} slide(s).");
        if (slideNumber == targetPosition)
            return;

        var slideIdToMove = slideIds[slideNumber - 1];
        slideIdToMove.Remove();

        var updatedSlideIds = slideIdList.Elements<SlideId>().ToList();
        var insertBeforeIndex = targetPosition - 1;

        if (insertBeforeIndex >= updatedSlideIds.Count)
            slideIdList.Append((SlideId)slideIdToMove.CloneNode(true));
        else
            slideIdList.InsertBefore((SlideId)slideIdToMove.CloneNode(true), updatedSlideIds[insertBeforeIndex]);

        presentationPart.Presentation.Save();
    }

    public void DeleteSlide(string filePath, int slideNumber)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var presentationPart = doc.PresentationPart!;
        var slideIdList = presentationPart.Presentation.SlideIdList
            ?? throw new InvalidOperationException("Presentation has no slides.");

        var slideIds = slideIdList.Elements<SlideId>().ToList();
        int count = slideIds.Count;

        if (count == 1)
            throw new InvalidOperationException("Cannot delete the only slide in a presentation.");
        if (slideNumber < 1 || slideNumber > count)
            throw new ArgumentOutOfRangeException(nameof(slideNumber), $"slideNumber {slideNumber} is out of range. Presentation has {count} slide(s).");

        var slideId = slideIds[slideNumber - 1];
        var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!);

        slideId.Remove();
        presentationPart.DeletePart(slidePart);

        presentationPart.Presentation.Save();
    }

    public void ReorderSlides(string filePath, int[] newOrder)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var presentationPart = doc.PresentationPart!;
        var slideIdList = presentationPart.Presentation.SlideIdList
            ?? throw new InvalidOperationException("Presentation has no slides.");

        var slideIds = slideIdList.Elements<SlideId>().ToList();
        int count = slideIds.Count;

        if (newOrder.Length != count)
            throw new ArgumentException($"newOrder must contain exactly {count} element(s), one per slide. Received {newOrder.Length}.", nameof(newOrder));

        var sorted = newOrder.OrderBy(n => n).ToList();
        for (int i = 0; i < sorted.Count; i++)
        {
            if (sorted[i] != i + 1)
                throw new ArgumentException($"newOrder must be a permutation of 1..{count}. Found invalid value {sorted[i]}.", nameof(newOrder));
        }

        var reordered = newOrder.Select(n => (SlideId)slideIds[n - 1].CloneNode(true)).ToList();

        foreach (var slideId in slideIds)
            slideId.Remove();

        foreach (var slideId in reordered)
            slideIdList.Append(slideId);

        presentationPart.Presentation.Save();
    }
}


