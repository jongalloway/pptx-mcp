using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Tests;

internal static class TemplateDeckHelper
{
    private static readonly byte[] SampleImageBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nZxQAAAAASUVORK5CYII=");

    public const string TitleBodyLayoutName = "Title and Body";
    public const string PictureCaptionLayoutName = "Picture Caption";

    public static void CreateTemplatePresentation(string filePath)
    {
        using var doc = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var titleBodyLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        var pictureCaptionLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

        titleBodyLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(CreateLayoutShapeTree(
                CreatePlaceholderShape(2U, "Layout Title", PlaceholderValues.Title, 0U, 457200, 274320, 8229600, 685800, "Click to add title"),
                CreatePlaceholderShape(3U, "Layout Body", PlaceholderValues.Body, 1U, 914400, 1600200, 7315200, 1371600, "Click to add text"),
                CreatePlaceholderShape(4U, "Layout Body 2", PlaceholderValues.Body, 2U, 914400, 3200400, 7315200, 914400, "Click to add text"))),
            new ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = SlideLayoutValues.Text
        };
        titleBodyLayoutPart.SlideLayout.CommonSlideData!.Name = TitleBodyLayoutName;

        pictureCaptionLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(CreateLayoutShapeTree(
                CreatePlaceholderShape(2U, "Picture Layout Title", PlaceholderValues.Title, 0U, 457200, 274320, 8229600, 685800, "Click to add title"),
                CreatePicturePlaceholder(3U, "Picture Placeholder", 1U, 914400, 1600200, 3657600, 2743200),
                CreatePlaceholderShape(4U, "Picture Caption Body", PlaceholderValues.Body, 2U, 5029200, 1600200, 2743200, 1143000, "Click to add text"))),
            new ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = SlideLayoutValues.Text
        };
        pictureCaptionLayoutPart.SlideLayout.CommonSlideData!.Name = PictureCaptionLayoutName;

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(CreateLayoutShapeTree()),
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
                new SlideLayoutId { Id = 2049U, RelationshipId = slideMasterPart.GetIdOfPart(titleBodyLayoutPart) },
                new SlideLayoutId { Id = 2050U, RelationshipId = slideMasterPart.GetIdOfPart(pictureCaptionLayoutPart) }));

        titleBodyLayoutPart.AddPart(slideMasterPart);
        pictureCaptionLayoutPart.AddPart(slideMasterPart);

        var firstSlidePart = presentationPart.AddNewPart<SlidePart>();
        firstSlidePart.AddPart(titleBodyLayoutPart);
        firstSlidePart.Slide = CreateSourceSlide(firstSlidePart);

        presentationPart.Presentation = new Presentation(
            new SlideIdList(
                new SlideId
                {
                    Id = 256U,
                    RelationshipId = presentationPart.GetIdOfPart(firstSlidePart)
                }),
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });

        presentationPart.Presentation.InsertAt(
            new SlideMasterIdList(
                new SlideMasterId
                {
                    Id = 2147483648U,
                    RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
                }),
            0);

        presentationPart.Presentation.Save();
    }

    private static Slide CreateSourceSlide(SlidePart slidePart)
    {
        var shapeTree = CreateLayoutShapeTree(
            CreatePlaceholderShape(2U, "Title 1", PlaceholderValues.Title, 0U, 457200, 274320, 8229600, 685800, "Quarterly Business Review"),
            CreatePlaceholderShape(3U, "Body 1", PlaceholderValues.Body, 1U, 914400, 1600200, 7315200, 1371600, "Revenue up 12%", "EMEA stable"),
            CreatePlaceholderShape(4U, "Body 2", PlaceholderValues.Body, 2U, 914400, 3200400, 7315200, 914400, "Follow-up items"));

        var imagePart = slidePart.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(SampleImageBytes))
            imagePart.FeedData(stream);

        var imageRelationshipId = slidePart.GetIdOfPart(imagePart);

        shapeTree.Append(CreatePicture(5U, imageRelationshipId, 5486400, 1600200, 2286000, 1828800));
        shapeTree.Append(CreatePicture(6U, imageRelationshipId, 5486400, 3657600, 1828800, 1371600));

        return new Slide(
            new CommonSlideData(shapeTree),
            new ColorMapOverride(new A.MasterColorMapping()));
    }

    private static ShapeTree CreateLayoutShapeTree(params OpenXmlElement[] shapes)
    {
        var shapeTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

        foreach (var shape in shapes)
            shapeTree.Append(shape);

        return shapeTree;
    }

    private static Shape CreatePlaceholderShape(uint shapeId, string name, PlaceholderValues placeholderType, uint placeholderIndex, long x, long y, long width, long height, params string[] paragraphs)
    {
        var textBody = new TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var paragraph in paragraphs.DefaultIfEmpty(string.Empty))
        {
            textBody.Append(new A.Paragraph(
                new A.Run(new A.Text(paragraph)),
                new A.EndParagraphRunProperties()));
        }

        return new Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape
                    {
                        Type = placeholderType,
                        Index = placeholderIndex
                    })),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = width, Cy = height })),
            textBody);
    }

    private static Picture CreatePicturePlaceholder(uint shapeId, string name, uint placeholderIndex, long x, long y, long width, long height) =>
        new(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape
                    {
                        Type = PlaceholderValues.Picture,
                        Index = placeholderIndex
                    })),
            new P.BlipFill(new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = width, Cy = height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

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
