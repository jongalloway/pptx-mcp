using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Tests;

internal static class TestPptxHelper
{
    private static readonly byte[] SampleImageBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nZxQAAAAASUVORK5CYII=");

    public static void CreateMinimalPresentation(string filePath, string? titleText = "Test Slide") =>
        CreatePresentation(filePath,
        [
            new TestSlideDefinition
            {
                TitleText = titleText
            }
        ]);

    public static void CreatePresentation(string filePath, IReadOnlyList<TestSlideDefinition> slides)
    {
        using var doc = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);

        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

        slideLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = SlideLayoutValues.Title,
        };
        slideLayoutPart.SlideLayout.CommonSlideData!.Name = "Title Slide";

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
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
            },
            new SlideLayoutIdList(
                new SlideLayoutId
                {
                    Id = 2049,
                    RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
                }));

        slideLayoutPart.AddPart(slideMasterPart);

        var slideIdList = new SlideIdList();
        uint nextSlideId = 256;

        foreach (var slideDefinition in slides)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(slideLayoutPart);
            slidePart.Slide = BuildSlide(slidePart, slideDefinition);

            if (!string.IsNullOrWhiteSpace(slideDefinition.SpeakerNotesText))
                AddSpeakerNotes(slidePart, slideDefinition.SpeakerNotesText!);

            slideIdList.Append(new SlideId
            {
                Id = nextSlideId++,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        presentationPart.Presentation = new Presentation(
            slideIdList,
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });

        var slideMasterIdList = new SlideMasterIdList(
            new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

        presentationPart.Presentation.InsertAt(slideMasterIdList, 0);
        presentationPart.Presentation.Save();
    }

    private static Slide BuildSlide(SlidePart slidePart, TestSlideDefinition slideDefinition)
    {
        var shapeTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

        uint nextShapeId = 2;

        if (!string.IsNullOrWhiteSpace(slideDefinition.TitleText))
        {
            shapeTree.Append(CreateTextShape(
                nextShapeId++,
                "Title 1",
                [new TestParagraphDefinition { Text = slideDefinition.TitleText }],
                PlaceholderValues.CenteredTitle,
                457200,
                274320,
                8229600,
                685800));
        }

        long currentY = 1371600;
        foreach (var textShape in slideDefinition.TextShapes)
        {
            var paragraphDefinitions = textShape.ParagraphDefinitions.Count > 0
                ? textShape.ParagraphDefinitions
                : textShape.Paragraphs.Select(paragraph => new TestParagraphDefinition { Text = paragraph }).ToList();

            long height = textShape.Height ?? Math.Max(685800, 342900L * Math.Max(1, paragraphDefinitions.Count));
            shapeTree.Append(CreateTextShape(
                nextShapeId,
                string.IsNullOrWhiteSpace(textShape.Name) ? $"Text {nextShapeId}" : textShape.Name!,
                paragraphDefinitions,
                textShape.PlaceholderType,
                textShape.X ?? 914400,
                textShape.Y ?? currentY,
                textShape.Width ?? 7315200,
                height));
            nextShapeId++;

            if (textShape.Y is null)
                currentY += height + 228600;
        }

        foreach (var table in slideDefinition.Tables)
        {
            long height = table.Height ?? 1371600;
            shapeTree.Append(CreateTable(
                nextShapeId,
                string.IsNullOrWhiteSpace(table.Name) ? $"Table {nextShapeId}" : table.Name!,
                table,
                table.X ?? 914400,
                table.Y ?? currentY,
                table.Width ?? 7315200,
                height));
            nextShapeId++;

            if (table.Y is null)
                currentY += height + 228600;
        }

        var slide = new Slide(
            new CommonSlideData(shapeTree),
            new ColorMapOverride(new A.MasterColorMapping()));

        if (slideDefinition.IncludeImage)
        {
            var imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using var imageStream = new MemoryStream(SampleImageBytes);
            imagePart.FeedData(imageStream);

            shapeTree.Append(CreatePicture(
                nextShapeId++,
                slidePart.GetIdOfPart(imagePart),
                914400,
                currentY,
                3657600,
                2743200));
        }

        return slide;
    }

    private static void AddSpeakerNotes(SlidePart slidePart, string speakerNotesText)
    {
        var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
        notesSlidePart.NotesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()),
                    CreateSpeakerNotesShape(speakerNotesText))),
            new ColorMapOverride(new A.MasterColorMapping()));
        notesSlidePart.AddPart(slidePart);
        notesSlidePart.NotesSlide.Save();
    }

    private static Shape CreateSpeakerNotesShape(string speakerNotesText) =>
        new(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2U, Name = "Notes Placeholder 2" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1U })),
            new ShapeProperties(),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(new A.Text(speakerNotesText)),
                    new A.EndParagraphRunProperties())));

    private static Shape CreateTextShape(
        uint shapeId,
        string name,
        IReadOnlyList<TestParagraphDefinition> paragraphs,
        PlaceholderValues? placeholderType,
        long x,
        long y,
        long width,
        long height)
    {
        var applicationProperties = placeholderType is null
            ? new ApplicationNonVisualDrawingProperties()
            : new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = placeholderType });

        var textBody = new TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var paragraph in paragraphs.DefaultIfEmpty(new TestParagraphDefinition()))
            textBody.Append(CreateParagraph(paragraph));

        return new Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new P.NonVisualShapeDrawingProperties(),
                applicationProperties),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = width, Cy = height })),
            textBody);
    }

    private static A.Paragraph CreateParagraph(TestParagraphDefinition paragraph)
    {
        var pptParagraph = new A.Paragraph();
        if (paragraph.IsBullet || paragraph.IsNumbered || paragraph.Level > 0)
        {
            var properties = new A.ParagraphProperties();
            if (paragraph.Level > 0)
                properties.Level = paragraph.Level;

            properties.Append(new A.CharacterBullet { Char = "•" });
            pptParagraph.Append(properties);
        }

        pptParagraph.Append(new A.Run(new A.Text(paragraph.Text ?? string.Empty)));
        pptParagraph.Append(new A.EndParagraphRunProperties());
        return pptParagraph;
    }

    private static P.GraphicFrame CreateTable(uint shapeId, string name, TestTableDefinition table, long x, long y, long width, long height)
    {
        var rows = table.Rows.Count == 0
            ? new List<List<string>> { new() { string.Empty } }
            : table.Rows.Select(row => row.Count == 0 ? new List<string> { string.Empty } : row.ToList()).ToList();
        var columnCount = rows.Max(row => row.Count);
        var rowHeight = Math.Max(342900L, height / rows.Count);
        var columnWidth = Math.Max(1L, width / columnCount);

        var drawingTable = new A.Table(new A.TableProperties { FirstRow = true, BandRow = true });
        var tableGrid = drawingTable.AppendChild(new A.TableGrid());
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            tableGrid.Append(new A.GridColumn { Width = columnWidth });

        foreach (var row in rows)
        {
            var normalizedRow = row.Concat(Enumerable.Repeat(string.Empty, columnCount - row.Count)).ToList();
            var tableRow = new A.TableRow { Height = rowHeight };
            foreach (var cellText in normalizedRow)
            {
                tableRow.Append(new A.TableCell(
                    new A.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(new A.Text(cellText ?? string.Empty)),
                            new A.EndParagraphRunProperties())),
                    new A.TableCellProperties()));
            }

            drawingTable.Append(tableRow);
        }

        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = width, Cy = height }),
            new A.Graphic(
                new A.GraphicData(drawingTable)
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
                }));
    }

    private static Picture CreatePicture(uint shapeId, string relationshipId, long x, long y, long width, long height) =>
        new(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Picture {shapeId}" },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = width, Cy = height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
}

public sealed class TestSlideDefinition
{
    public string? TitleText { get; init; }

    public string? SpeakerNotesText { get; init; }

    public IReadOnlyList<TestTextShapeDefinition> TextShapes { get; init; } = [];

    public IReadOnlyList<TestTableDefinition> Tables { get; init; } = [];

    public bool IncludeImage { get; init; }
}

public sealed class TestTextShapeDefinition
{
    public string? Name { get; init; }

    public IReadOnlyList<string> Paragraphs { get; init; } = [];

    public IReadOnlyList<TestParagraphDefinition> ParagraphDefinitions { get; init; } = [];

    public PlaceholderValues? PlaceholderType { get; init; }

    public long? X { get; init; }

    public long? Y { get; init; }

    public long? Width { get; init; }

    public long? Height { get; init; }
}

public sealed class TestParagraphDefinition
{
    public string? Text { get; init; }

    public int Level { get; init; }

    public bool IsBullet { get; init; }

    public bool IsNumbered { get; init; }
}

public sealed class TestTableDefinition
{
    public string? Name { get; init; }

    public IReadOnlyList<IReadOnlyList<string>> Rows { get; init; } = [];

    public long? X { get; init; }

    public long? Y { get; init; }

    public long? Width { get; init; }

    public long? Height { get; init; }
}
