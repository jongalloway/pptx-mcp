using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests.Services;

[Trait("Category", "Unit")]
public class ImageReplaceTests : PptxTestBase
{
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nZxQAAAAASUVORK5CYII=");

    // Minimal valid JPEG (1x1 pixel)
    private static readonly byte[] JpegBytes = Convert.FromBase64String(
        "/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYI4Q/RFhHRUYnJCkqNTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqCg4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ufo6ery8/T19vf4+fr/2gAMAwEAAhEDEQA/AD8A/9k=");

    // Minimal SVG
    private static readonly byte[] SvgBytes = System.Text.Encoding.UTF8.GetBytes(
        "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1\" height=\"1\"><rect width=\"1\" height=\"1\" fill=\"red\"/></svg>");

    private string CreateTrackedPath(string? extension = ".pptx")
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + extension);
        TrackTempFile(path);
        return path;
    }

    private string CreatePptxWithPicture(string pictureName = "Picture 2", int pictureCount = 1)
    {
        var pptxPath = CreateTrackedPath();
        using var doc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Create(pptxPath, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);

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
        { Type = SlideLayoutValues.Title };
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
            new SlideLayoutIdList(new SlideLayoutId
            {
                Id = 2049,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            }));

        slideLayoutPart.AddPart(slideMasterPart);

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.AddPart(slideLayoutPart);

        var shapeTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

        uint nextId = 2;
        for (int i = 0; i < pictureCount; i++)
        {
            var name = pictureCount == 1 ? pictureName : $"{pictureName} {i}";
            var imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using var imageStream = new MemoryStream(PngBytes);
            imagePart.FeedData(imageStream);

            shapeTree.Append(TestPptxHelper.CreatePicture(
                nextId, slidePart.GetIdOfPart(imagePart),
                914400, 914400, 3657600, 2743200, name));
            nextId++;
        }

        slidePart.Slide = new Slide(
            new CommonSlideData(shapeTree),
            new ColorMapOverride(new A.MasterColorMapping()));

        var slideIdList = new SlideIdList(new SlideId
        {
            Id = 256,
            RelationshipId = presentationPart.GetIdOfPart(slidePart)
        });

        presentationPart.Presentation = new Presentation(
            slideIdList,
            new SlideSize { Cx = (int)Emu.Inches10, Cy = (int)Emu.Inches7_5, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = (int)Emu.Inches7_5, Cy = (int)Emu.Inches10 });

        presentationPart.Presentation.InsertAt(
            new SlideMasterIdList(new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            }), 0);

        presentationPart.Presentation.Save();
        return pptxPath;
    }

    private string CreateTempImage(byte[] bytes, string extension = ".png")
    {
        var path = CreateTrackedPath(extension);
        File.WriteAllBytes(path, bytes);
        return path;
    }

    #region Replacement by name

    [Fact]
    public void ReplaceImage_ByName_Succeeds()
    {
        var pptxPath = CreatePptxWithPicture("Logo");
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        var result = Service.ReplaceImage(pptxPath, 1, "Logo", null, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal("Logo", result.ShapeName);
        Assert.Equal("shapeName", result.MatchedBy);
        Assert.Equal("image/jpeg", result.NewImageContentType);
        Assert.Equal("image/png", result.PreviousImageContentType);
    }

    [Fact]
    public void ReplaceImage_ByName_CaseInsensitive()
    {
        var pptxPath = CreatePptxWithPicture("Company Logo");
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 1, "company logo", null, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal("Company Logo", result.ShapeName);
    }

    [Fact]
    public void ReplaceImage_ByName_NotFound_ReportsAvailable()
    {
        var pptxPath = CreatePptxWithPicture("Logo");
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 1, "NonExistent", null, imagePath, null);

        Assert.False(result.Success);
        Assert.Contains("Logo", result.Message);
    }

    #endregion

    #region Replacement by index

    [Fact]
    public void ReplaceImage_ByIndex_Succeeds()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        var result = Service.ReplaceImage(pptxPath, 1, null, 0, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal("Photo", result.ShapeName);
        Assert.Equal("shapeIndex", result.MatchedBy);
    }

    [Fact]
    public void ReplaceImage_ByIndex_OutOfRange()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 1, null, 5, imagePath, null);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message);
    }

    #endregion

    #region Alt text

    [Fact]
    public void ReplaceImage_SetsAltText()
    {
        var pptxPath = CreatePptxWithPicture("Hero Image");
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 1, "Hero Image", null, imagePath, "A scenic mountain view");

        Assert.True(result.Success);
        Assert.Equal("A scenic mountain view", result.AltText);

        // Verify alt text was persisted
        using var doc = PresentationDocument.Open(pptxPath, false);
        var slidePart = (SlidePart)doc.PresentationPart!.GetPartById(
            doc.PresentationPart.Presentation.SlideIdList!.Elements<SlideId>().First().RelationshipId!.Value!);
        var picture = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Picture>().First();
        Assert.Equal("A scenic mountain view", picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value);
    }

    [Fact]
    public void ReplaceImage_NullAltText_PreservesExisting()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var imagePath = CreateTempImage(PngBytes, ".png");

        // Alt text not set initially, should remain null
        var result = Service.ReplaceImage(pptxPath, 1, "Photo", null, imagePath, null);

        Assert.True(result.Success);
        Assert.Null(result.AltText);
    }

    #endregion

    #region Format validation

    [Fact]
    public void ReplaceImage_Png_Supported()
    {
        var pptxPath = CreatePptxWithPicture();
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 1, null, 0, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal("image/png", result.NewImageContentType);
    }

    [Fact]
    public void ReplaceImage_Jpeg_Supported()
    {
        var pptxPath = CreatePptxWithPicture();
        var imagePath = CreateTempImage(JpegBytes, ".jpeg");

        var result = Service.ReplaceImage(pptxPath, 1, null, 0, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal("image/jpeg", result.NewImageContentType);
    }

    [Fact]
    public void ReplaceImage_Jpg_Supported()
    {
        var pptxPath = CreatePptxWithPicture();
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        var result = Service.ReplaceImage(pptxPath, 1, null, 0, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal("image/jpeg", result.NewImageContentType);
    }

    [Fact]
    public void ReplaceImage_Svg_Supported()
    {
        var pptxPath = CreatePptxWithPicture();
        var imagePath = CreateTempImage(SvgBytes, ".svg");

        var result = Service.ReplaceImage(pptxPath, 1, null, 0, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal("image/svg+xml", result.NewImageContentType);
    }

    [Fact]
    public void ReplaceImage_UnsupportedFormat_ReturnsError()
    {
        var pptxPath = CreatePptxWithPicture();
        var imagePath = CreateTempImage(PngBytes, ".tiff");

        var result = Service.ReplaceImage(pptxPath, 1, null, 0, imagePath, null);

        Assert.False(result.Success);
        Assert.Contains("Unsupported image format", result.Message);
    }

    #endregion

    #region Edge cases

    [Fact]
    public void ReplaceImage_NoSlides_ReturnsError()
    {
        var pptxPath = CreateTrackedPath();
        using (var doc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Create(pptxPath, DocumentFormat.OpenXml.PresentationDocumentType.Presentation))
        {
            var pp = doc.AddPresentationPart();
            pp.Presentation = new Presentation(
                new SlideIdList(),
                new SlideSize { Cx = (int)Emu.Inches10, Cy = (int)Emu.Inches7_5 },
                new NotesSize { Cx = (int)Emu.Inches7_5, Cy = (int)Emu.Inches10 });
            pp.Presentation.Save();
        }
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 1, null, 0, imagePath, null);

        Assert.False(result.Success);
        Assert.Contains("no slides", result.Message);
    }

    [Fact]
    public void ReplaceImage_SlideNumberOutOfRange_ReturnsError()
    {
        var pptxPath = CreatePptxWithPicture();
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 99, null, 0, imagePath, null);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message);
    }

    [Fact]
    public void ReplaceImage_SlideNumberZero_ReturnsError()
    {
        var pptxPath = CreatePptxWithPicture();
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 0, null, 0, imagePath, null);

        Assert.False(result.Success);
        Assert.Contains("must be 1 or greater", result.Message);
    }

    [Fact]
    public void ReplaceImage_NoShapeNameOrIndex_ReturnsError()
    {
        var pptxPath = CreatePptxWithPicture();
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 1, null, null, imagePath, null);

        Assert.False(result.Success);
        Assert.Contains("Provide either", result.Message);
    }

    [Fact]
    public void ReplaceImage_NoPicturesOnSlide_ReturnsError()
    {
        // Create a presentation with only text shapes, no pictures
        var pptxPath = CreateTrackedPath();
        TestPptxHelper.CreateMinimalPresentation(pptxPath, "Text Only Slide");
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 1, null, 0, imagePath, null);

        Assert.False(result.Success);
        Assert.Contains("does not contain any picture shapes", result.Message);
    }

    [Fact]
    public void ReplaceImage_MultiplePictures_ByIndex()
    {
        var pptxPath = CreatePptxWithPicture("Photo", pictureCount: 3);
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        var result = Service.ReplaceImage(pptxPath, 1, null, 1, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal("Photo 1", result.ShapeName);
        Assert.Equal("shapeIndex", result.MatchedBy);
    }

    [Fact]
    public void ReplaceImage_NegativeIndex_ReturnsError()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 1, null, -1, imagePath, null);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message);
    }

    [Fact]
    public void ReplaceImage_DuplicateNames_ReturnsAmbiguityError()
    {
        // Create a PPTX with two pictures sharing the same name
        var pptxPath = CreateTrackedPath();
        using (var doc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Create(pptxPath, DocumentFormat.OpenXml.PresentationDocumentType.Presentation))
        {
            var presentationPart = doc.AddPresentationPart();
            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

            slideLayoutPart.SlideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
                new ColorMapOverride(new A.MasterColorMapping())) { Type = SlideLayoutValues.Title };
            slideLayoutPart.SlideLayout.CommonSlideData!.Name = "Title Slide";

            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
                new P.ColorMap
                {
                    Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1,
                    Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2,
                    Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2,
                    Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4,
                    Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6,
                    Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
                },
                new SlideLayoutIdList(new SlideLayoutId { Id = 2049, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }));
            slideLayoutPart.AddPart(slideMasterPart);

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(slideLayoutPart);

            var imgPart = slidePart.AddImagePart(ImagePartType.Png);
            using (var ms = new MemoryStream(PngBytes)) imgPart.FeedData(ms);
            var relId = slidePart.GetIdOfPart(imgPart);

            var shapeTree = new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()));

            // Two pictures with same name
            for (uint id = 2; id <= 3; id++)
            {
                shapeTree.Append(TestPptxHelper.CreatePicture(
                    id, relId, 914400, 914400, 3657600, 2743200, "Logo"));
            }

            slidePart.Slide = new Slide(new CommonSlideData(shapeTree), new ColorMapOverride(new A.MasterColorMapping()));
            presentationPart.Presentation = new Presentation(
                new SlideIdList(new SlideId { Id = 256, RelationshipId = presentationPart.GetIdOfPart(slidePart) }),
                new SlideSize { Cx = (int)Emu.Inches10, Cy = (int)Emu.Inches7_5, Type = SlideSizeValues.Screen4x3 },
                new NotesSize { Cx = (int)Emu.Inches7_5, Cy = (int)Emu.Inches10 });
            presentationPart.Presentation.InsertAt(
                new SlideMasterIdList(new SlideMasterId { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }), 0);
            presentationPart.Presentation.Save();
        }

        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 1, "Logo", null, imagePath, null);

        Assert.False(result.Success);
        Assert.Contains("Multiple picture shapes named", result.Message);
        Assert.Contains("shapeIndex", result.Message);
    }

    [Fact]
    public void ReplaceImage_EmptyShapeName_ReturnsError()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var imagePath = CreateTempImage(PngBytes, ".png");

        var result = Service.ReplaceImage(pptxPath, 1, "   ", null, imagePath, null);

        Assert.False(result.Success);
        Assert.Contains("Provide either", result.Message);
    }

    [Fact]
    public void ReplaceImage_NameFallbackToIndex_UsesShapeIndexFallback()
    {
        var pptxPath = CreatePptxWithPicture("Photo", pictureCount: 2);
        var imagePath = CreateTempImage(PngBytes, ".png");

        // Name doesn't match, but index is valid → shapeIndexFallback
        var result = Service.ReplaceImage(pptxPath, 1, "NonExistent", 0, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal("shapeIndexFallback", result.MatchedBy);
    }

    #endregion

    #region Image bytes and relationship verification

    [Fact]
    public void ReplaceImage_ReplacesActualImageBytes()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var imagePath = CreateTempImage(JpegBytes, ".jpg");
        var expectedBytes = File.ReadAllBytes(imagePath);

        Service.ReplaceImage(pptxPath, 1, "Photo", null, imagePath, null);

        using var doc = PresentationDocument.Open(pptxPath, false);
        var slidePart = (SlidePart)doc.PresentationPart!.GetPartById(
            doc.PresentationPart.Presentation.SlideIdList!.Elements<SlideId>().First().RelationshipId!.Value!);
        var blip = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Picture>().First()
            .GetFirstChild<P.BlipFill>()!.GetFirstChild<A.Blip>()!;
        var part = slidePart.GetPartById(blip.Embed!.Value!);

        using var stream = part.GetStream();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        Assert.Equal(expectedBytes, ms.ToArray());
    }

    [Fact]
    public void ReplaceImage_OtherPicturesPreserved()
    {
        var pptxPath = CreatePptxWithPicture("Photo", pictureCount: 3);
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        // Replace the middle picture (index 1)
        Service.ReplaceImage(pptxPath, 1, null, 1, imagePath, null);

        // All 3 pictures should still exist
        using var doc = PresentationDocument.Open(pptxPath, false);
        var slidePart = (SlidePart)doc.PresentationPart!.GetPartById(
            doc.PresentationPart.Presentation.SlideIdList!.Elements<SlideId>().First().RelationshipId!.Value!);
        var pictures = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Picture>().ToList();
        Assert.Equal(3, pictures.Count);
    }

    [Fact]
    public void ReplaceImage_ShapeIdPreserved()
    {
        var pptxPath = CreatePptxWithPicture("Photo");

        // Capture original shape ID
        uint? originalId;
        using (var doc = PresentationDocument.Open(pptxPath, false))
        {
            var slidePart = (SlidePart)doc.PresentationPart!.GetPartById(
                doc.PresentationPart.Presentation.SlideIdList!.Elements<SlideId>().First().RelationshipId!.Value!);
            originalId = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Picture>().First()
                .NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
        }

        var imagePath = CreateTempImage(JpegBytes, ".jpg");
        Service.ReplaceImage(pptxPath, 1, "Photo", null, imagePath, null);

        using var doc2 = PresentationDocument.Open(pptxPath, false);
        var slidePart2 = (SlidePart)doc2.PresentationPart!.GetPartById(
            doc2.PresentationPart.Presentation.SlideIdList!.Elements<SlideId>().First().RelationshipId!.Value!);
        var newId = slidePart2.Slide.CommonSlideData!.ShapeTree!.Elements<Picture>().First()
            .NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;

        Assert.Equal(originalId, newId);
    }

    [Fact]
    public void ReplaceImage_CrossFormatReplacement_TracksPreviousContentType()
    {
        var pptxPath = CreatePptxWithPicture("Photo"); // starts as PNG
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        var result = Service.ReplaceImage(pptxPath, 1, "Photo", null, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal("image/png", result.PreviousImageContentType);
        Assert.Equal("image/jpeg", result.NewImageContentType);
    }

    [Fact]
    public void ReplaceImage_SuccessiveReplacements_AllSucceed()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var firstImage = CreateTempImage(JpegBytes, ".jpg");
        var secondImage = CreateTempImage(PngBytes, ".png");

        var r1 = Service.ReplaceImage(pptxPath, 1, "Photo", null, firstImage, "First");
        Assert.True(r1.Success);

        var r2 = Service.ReplaceImage(pptxPath, 1, "Photo", null, secondImage, "Second");
        Assert.True(r2.Success);
        Assert.Equal("image/jpeg", r2.PreviousImageContentType);
        Assert.Equal("image/png", r2.NewImageContentType);
        Assert.Equal("Second", r2.AltText);
    }

    [Fact]
    public void ReplaceImage_OnSecondSlide_TargetsCorrectSlide()
    {
        // Create two-slide deck with images
        var pptxPath = CreateTrackedPath();
        TestPptxHelper.CreatePresentation(pptxPath,
        [
            new TestSlideDefinition { TitleText = "Slide 1", IncludeImage = true },
            new TestSlideDefinition { TitleText = "Slide 2", IncludeImage = true }
        ]);
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        var result = Service.ReplaceImage(pptxPath, 2, null, 0, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal(2, result.SlideNumber);
    }

    [Fact]
    public void ReplaceImage_WithTemplateDeckPictures_Succeeds()
    {
        var pptxPath = CreateTrackedPath();
        TemplateDeckHelper.CreateTemplatePresentation(pptxPath);
        var imagePath = CreateTempImage(PngBytes, ".png");

        // TemplateDeckHelper's source slide has Picture shapes at indices 0 and 1
        var result = Service.ReplaceImage(pptxPath, 1, null, 0, imagePath, null);

        Assert.True(result.Success);
        Assert.Equal("shapeIndex", result.MatchedBy);
    }

    [Fact]
    public void ReplaceImage_ValidationErrorCountUnchanged()
    {
        var pptxPath = CreatePptxWithPicture("Logo");
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        using var baselineDoc = PresentationDocument.Open(pptxPath, false);
        var validator = new OpenXmlValidator();
        var baselineCount = validator.Validate(baselineDoc).Count();
        baselineDoc.Dispose();

        Service.ReplaceImage(pptxPath, 1, "Logo", null, imagePath, "Company Logo");

        using var postDoc = PresentationDocument.Open(pptxPath, false);
        var postCount = validator.Validate(postDoc).Count();

        Assert.Equal(baselineCount, postCount);
    }

    #endregion

    #region PowerPoint compatibility

    [Fact]
    public void ReplaceImage_PassesOpenXmlValidation()
    {
        var pptxPath = CreatePptxWithPicture("Logo");
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        Service.ReplaceImage(pptxPath, 1, "Logo", null, imagePath, "Company logo");

        // Collect baseline validation errors from fixture (SlideMaster warnings are benign)
        var baselinePath = CreatePptxWithPicture("Baseline");
        using var baselineDoc = PresentationDocument.Open(baselinePath, false);
        var validator = new OpenXmlValidator();
        var baselineErrors = validator.Validate(baselineDoc)
            .Select(e => e.Description)
            .ToHashSet();

        using var doc = PresentationDocument.Open(pptxPath, false);
        var errors = validator.Validate(doc)
            .Where(e => !baselineErrors.Contains(e.Description))
            .ToList();

        Assert.Empty(errors);
    }

    [Fact]
    public void ReplaceImage_PreservesShapeGeometry()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        Service.ReplaceImage(pptxPath, 1, "Photo", null, imagePath, null);

        using var doc = PresentationDocument.Open(pptxPath, false);
        var slidePart = (SlidePart)doc.PresentationPart!.GetPartById(
            doc.PresentationPart.Presentation.SlideIdList!.Elements<SlideId>().First().RelationshipId!.Value!);
        var picture = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Picture>().First();

        // Verify shape properties (position/size) are preserved
        var transform = picture.ShapeProperties?.GetFirstChild<A.Transform2D>();
        Assert.NotNull(transform);
        Assert.Equal(Emu.OneInch, transform!.Offset!.X!.Value);
        Assert.Equal(Emu.OneInch, transform.Offset!.Y!.Value);
        Assert.Equal(Emu.Inches4, transform.Extents!.Cx!.Value);
        Assert.Equal(Emu.Inches3, transform.Extents!.Cy!.Value);
    }

    [Fact]
    public void ReplaceImage_NewBlipPointsToReplacementImage()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        Service.ReplaceImage(pptxPath, 1, "Photo", null, imagePath, null);

        using var doc = PresentationDocument.Open(pptxPath, false);
        var slidePart = (SlidePart)doc.PresentationPart!.GetPartById(
            doc.PresentationPart.Presentation.SlideIdList!.Elements<SlideId>().First().RelationshipId!.Value!);
        var picture = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Picture>().First();
        var blip = picture.GetFirstChild<P.BlipFill>()?.GetFirstChild<A.Blip>();

        Assert.NotNull(blip);
        Assert.NotNull(blip!.Embed?.Value);

        // Verify the referenced part is JPEG
        var imagePart = slidePart.GetPartById(blip.Embed!.Value!);
        Assert.Equal("image/jpeg", imagePart.ContentType);
    }

    #endregion
}
