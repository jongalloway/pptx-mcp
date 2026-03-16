namespace PptxMcp.Tests.Services;

public class PresentationServiceTests : IDisposable
{
    private readonly PresentationService _service = new();
    private readonly List<string> _tempFiles = new();

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            if (File.Exists(f)) File.Delete(f);
    }

    private string CreateTempPptx(string? titleText = "Test Slide")
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        _tempFiles.Add(path);
        TestPptxHelper.CreateMinimalPresentation(path, titleText);
        return path;
    }

    [Fact]
    public void GetSlides_ReturnsCorrectCount()
    {
        var path = CreateTempPptx();
        var slides = _service.GetSlides(path);
        Assert.Single(slides);
    }

    [Fact]
    public void GetSlides_ReturnsCorrectTitle()
    {
        var path = CreateTempPptx("Hello World");
        var slides = _service.GetSlides(path);
        Assert.Equal("Hello World", slides[0].Title);
    }

    [Fact]
    public void GetSlides_SlideHasCorrectIndex()
    {
        var path = CreateTempPptx();
        var slides = _service.GetSlides(path);
        Assert.Equal(0, slides[0].Index);
    }

    [Fact]
    public void GetLayouts_ReturnsLayouts()
    {
        var path = CreateTempPptx();
        var layouts = _service.GetLayouts(path);
        Assert.NotEmpty(layouts);
    }

    [Fact]
    public void GetLayouts_LayoutHasName()
    {
        var path = CreateTempPptx();
        var layouts = _service.GetLayouts(path);
        Assert.All(layouts, l => Assert.NotNull(l.Name));
    }

    [Fact]
    public void AddSlide_IncreasesSlideCount()
    {
        var path = CreateTempPptx();
        var before = _service.GetSlides(path);
        _service.AddSlide(path, null);
        var after = _service.GetSlides(path);
        Assert.Equal(before.Count + 1, after.Count);
    }

    [Fact]
    public void AddSlide_ReturnsNewSlideIndex()
    {
        var path = CreateTempPptx();
        var newIndex = _service.AddSlide(path, null);
        Assert.Equal(1, newIndex);
    }

    [Fact]
    public void UpdateTextPlaceholder_ChangesTextContent()
    {
        var path = CreateTempPptx("Original Title");
        _service.UpdateTextPlaceholder(path, 0, 0, "Updated Title");
        var slides = _service.GetSlides(path);
        Assert.Equal("Updated Title", slides[0].Title);
    }

    [Fact]
    public void GetSlideXml_ReturnsXmlString()
    {
        var path = CreateTempPptx();
        var xml = _service.GetSlideXml(path, 0);
        Assert.NotNull(xml);
        Assert.Contains("sld", xml);
    }

    [Fact]
    public void GetSlideXml_OutOfRange_ThrowsException()
    {
        var path = CreateTempPptx();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _service.GetSlideXml(path, 99));
    }

    [Fact]
    public void GetSlideContent_ReturnsSlideIndex()
    {
        var path = CreateTempPptx();
        var content = _service.GetSlideContent(path, 0);
        Assert.Equal(0, content.SlideIndex);
    }

    [Fact]
    public void GetSlideContent_ReturnsSlideDimensions()
    {
        var path = CreateTempPptx();
        var content = _service.GetSlideContent(path, 0);
        Assert.True(content.SlideWidthEmu > 0);
        Assert.True(content.SlideHeightEmu > 0);
    }

    [Fact]
    public void GetSlideContent_ReturnsShapes()
    {
        var path = CreateTempPptx();
        var content = _service.GetSlideContent(path, 0);
        Assert.NotEmpty(content.Shapes);
    }

    [Fact]
    public void GetSlideContent_TitleShapeHasText()
    {
        var path = CreateTempPptx("My Title");
        var content = _service.GetSlideContent(path, 0);
        var titleShape = content.Shapes.FirstOrDefault(s => s.IsPlaceholder);
        Assert.NotNull(titleShape);
        Assert.Equal("My Title", titleShape.Text);
    }

    [Fact]
    public void GetSlideContent_TitleShapeIsTextType()
    {
        var path = CreateTempPptx();
        var content = _service.GetSlideContent(path, 0);
        var titleShape = content.Shapes.FirstOrDefault(s => s.IsPlaceholder);
        Assert.NotNull(titleShape);
        Assert.Equal("Text", titleShape.ShapeType);
    }

    [Fact]
    public void GetSlideContent_TitleShapeHasPlaceholderType()
    {
        var path = CreateTempPptx();
        var content = _service.GetSlideContent(path, 0);
        var titleShape = content.Shapes.FirstOrDefault(s => s.IsPlaceholder);
        Assert.NotNull(titleShape);
        Assert.NotNull(titleShape.PlaceholderType);
    }

    [Fact]
    public void GetSlideContent_OutOfRange_ThrowsException()
    {
        var path = CreateTempPptx();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _service.GetSlideContent(path, 99));
    }
}
