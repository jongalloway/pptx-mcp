using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Tests.Tools;

public class ImageReplaceToolTests : IDisposable
{
    private readonly PresentationService _service = new();
    private readonly PptxTools _tools;
    private readonly List<string> _tempFiles = [];

    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nZxQAAAAASUVORK5CYII=");

    private static readonly byte[] JpegBytes = Convert.FromBase64String(
        "/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYI4Q/RFhHRUYnJCkqNTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqCg4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ufo6ery8/T19vf4+fr/2gAMAwEAAhEDEQA/AD8A/9k=");

    public ImageReplaceToolTests()
    {
        _tools = new PptxTools(_service);
    }

    public void Dispose()
    {
        foreach (var file in _tempFiles)
            if (File.Exists(file)) File.Delete(file);
    }

    private string TrackTempFile(string extension = ".pptx")
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + extension);
        _tempFiles.Add(path);
        return path;
    }

    private string CreatePptxWithPicture(string pictureName = "Photo")
    {
        var pptxPath = TrackTempFile();
        using var doc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Create(
            pptxPath, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);

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

        var imagePart = slidePart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(PngBytes)) imagePart.FeedData(ms);

        var shapeTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

        shapeTree.Append(TestPptxHelper.CreatePicture(
            2, slidePart.GetIdOfPart(imagePart),
            914400, 914400, 3657600, 2743200, pictureName));

        slidePart.Slide = new Slide(new CommonSlideData(shapeTree), new ColorMapOverride(new A.MasterColorMapping()));

        presentationPart.Presentation = new Presentation(
            new SlideIdList(new SlideId { Id = 256, RelationshipId = presentationPart.GetIdOfPart(slidePart) }),
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });
        presentationPart.Presentation.InsertAt(
            new SlideMasterIdList(new SlideMasterId { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }), 0);
        presentationPart.Presentation.Save();

        return pptxPath;
    }

    private string CreateTempImage(byte[] bytes, string extension = ".png")
    {
        var path = TrackTempFile(extension);
        File.WriteAllBytes(path, bytes);
        return path;
    }

    #region JSON output format

    [Fact]
    public async Task pptx_replace_image_Success_ReturnsStructuredJson()
    {
        var pptxPath = CreatePptxWithPicture("Hero Image");
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        var json = await _tools.pptx_replace_image(pptxPath, 1, "Hero Image", null, imagePath, "Alt text");

        var result = JsonSerializer.Deserialize<ImageReplaceResult>(json);
        Assert.NotNull(result);
        Assert.True(result!.Success);
        Assert.Equal(1, result.SlideNumber);
        Assert.Equal("Hero Image", result.ShapeName);
        Assert.Equal("shapeName", result.MatchedBy);
        Assert.Equal("image/jpeg", result.NewImageContentType);
        Assert.Equal("Alt text", result.AltText);
        Assert.Contains("Hero Image", result.Message);
    }

    [Fact]
    public async Task pptx_replace_image_Failure_ReturnsStructuredJson()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var imagePath = CreateTempImage(PngBytes, ".png");

        var json = await _tools.pptx_replace_image(pptxPath, 1, "NonExistent", null, imagePath, null);

        var result = JsonSerializer.Deserialize<ImageReplaceResult>(json);
        Assert.NotNull(result);
        Assert.False(result!.Success);
        Assert.Equal(1, result.SlideNumber);
        Assert.Contains("Photo", result.Message);
    }

    #endregion

    #region File not found error messages

    [Fact]
    public async Task pptx_replace_image_PptxNotFound_ReturnsJsonError()
    {
        var fakePptx = Path.Join(Path.GetTempPath(), "nonexistent_file.pptx");
        var imagePath = CreateTempImage(PngBytes, ".png");

        var json = await _tools.pptx_replace_image(fakePptx, 1, "Photo", null, imagePath, null);

        var result = JsonSerializer.Deserialize<ImageReplaceResult>(json);
        Assert.NotNull(result);
        Assert.False(result!.Success);
        Assert.Contains("File not found", result.Message);
        Assert.Contains("nonexistent_file.pptx", result.Message);
    }

    [Fact]
    public async Task pptx_replace_image_ImageNotFound_ReturnsJsonError()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var fakeImage = Path.Join(Path.GetTempPath(), "nonexistent_image.png");

        var json = await _tools.pptx_replace_image(pptxPath, 1, "Photo", null, fakeImage, null);

        var result = JsonSerializer.Deserialize<ImageReplaceResult>(json);
        Assert.NotNull(result);
        Assert.False(result!.Success);
        Assert.Contains("Image file not found", result.Message);
        Assert.Contains("nonexistent_image.png", result.Message);
    }

    #endregion

    #region Parameter validation at tool level

    [Fact]
    public async Task pptx_replace_image_BothFilesMissing_ReportsFileNotFound()
    {
        var fakePptx = Path.Join(Path.GetTempPath(), "no_such.pptx");

        var json = await _tools.pptx_replace_image(fakePptx, 1, "Photo", null, "no_such.png", null);

        var result = JsonSerializer.Deserialize<ImageReplaceResult>(json);
        Assert.NotNull(result);
        Assert.False(result!.Success);
        // File check happens before image check
        Assert.Contains("File not found", result.Message);
    }

    [Fact]
    public async Task pptx_replace_image_UnsupportedFormat_ReturnsJsonError()
    {
        var pptxPath = CreatePptxWithPicture("Photo");
        var imagePath = CreateTempImage(PngBytes, ".tiff");

        var json = await _tools.pptx_replace_image(pptxPath, 1, "Photo", null, imagePath, null);

        var result = JsonSerializer.Deserialize<ImageReplaceResult>(json);
        Assert.NotNull(result);
        Assert.False(result!.Success);
        Assert.Contains("Unsupported image format", result.Message);
    }

    [Fact]
    public async Task pptx_replace_image_AllJsonFieldsPresent()
    {
        var pptxPath = CreatePptxWithPicture("Logo");
        var imagePath = CreateTempImage(JpegBytes, ".jpg");

        var json = await _tools.pptx_replace_image(pptxPath, 1, "Logo", null, imagePath, "Company logo");

        // Verify the JSON contains all expected fields
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        Assert.True(root.TryGetProperty("Success", out _));
        Assert.True(root.TryGetProperty("SlideNumber", out _));
        Assert.True(root.TryGetProperty("ShapeName", out _));
        Assert.True(root.TryGetProperty("MatchedBy", out _));
        Assert.True(root.TryGetProperty("PreviousImageContentType", out _));
        Assert.True(root.TryGetProperty("NewImageContentType", out _));
        Assert.True(root.TryGetProperty("AltText", out _));
        Assert.True(root.TryGetProperty("Message", out _));
    }

    [Fact]
    public async Task pptx_replace_image_JsonIsIndented()
    {
        var pptxPath = CreatePptxWithPicture("Logo");
        var imagePath = CreateTempImage(PngBytes, ".png");

        var json = await _tools.pptx_replace_image(pptxPath, 1, "Logo", null, imagePath, null);

        // Indented JSON has newlines
        Assert.Contains("\n", json);
        Assert.Contains("  ", json);
    }

    #endregion
}
