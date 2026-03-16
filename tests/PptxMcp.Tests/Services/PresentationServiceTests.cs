using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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
        var path = Path.GetTempFileName() + ".pptx";
        _tempFiles.Add(path);
        TestPptxHelper.CreateMinimalPresentation(path, titleText);
        return path;
    }

    [Fact]
    public async Task GetSlidesAsync_ReturnsCorrectCount()
    {
        var path = CreateTempPptx();
        var slides = await _service.GetSlidesAsync(path);
        Assert.Single(slides);
    }

    [Fact]
    public async Task GetSlidesAsync_ReturnsCorrectTitle()
    {
        var path = CreateTempPptx("Hello World");
        var slides = await _service.GetSlidesAsync(path);
        Assert.Equal("Hello World", slides[0].Title);
    }

    [Fact]
    public async Task GetSlidesAsync_SlideHasCorrectIndex()
    {
        var path = CreateTempPptx();
        var slides = await _service.GetSlidesAsync(path);
        Assert.Equal(0, slides[0].Index);
    }

    [Fact]
    public async Task GetLayoutsAsync_ReturnsLayouts()
    {
        var path = CreateTempPptx();
        var layouts = await _service.GetLayoutsAsync(path);
        Assert.NotEmpty(layouts);
    }

    [Fact]
    public async Task GetLayoutsAsync_LayoutHasName()
    {
        var path = CreateTempPptx();
        var layouts = await _service.GetLayoutsAsync(path);
        Assert.All(layouts, l => Assert.NotNull(l.Name));
    }

    [Fact]
    public async Task AddSlideAsync_IncreasesSlideCount()
    {
        var path = CreateTempPptx();
        var before = await _service.GetSlidesAsync(path);
        await _service.AddSlideAsync(path, null);
        var after = await _service.GetSlidesAsync(path);
        Assert.Equal(before.Count + 1, after.Count);
    }

    [Fact]
    public async Task AddSlideAsync_ReturnsNewSlideIndex()
    {
        var path = CreateTempPptx();
        var newIndex = await _service.AddSlideAsync(path, null);
        Assert.Equal(1, newIndex);
    }

    [Fact]
    public async Task UpdateTextPlaceholderAsync_ChangesTextContent()
    {
        var path = CreateTempPptx("Original Title");
        await _service.UpdateTextPlaceholderAsync(path, 0, 0, "Updated Title");
        var slides = await _service.GetSlidesAsync(path);
        Assert.Equal("Updated Title", slides[0].Title);
    }

    [Fact]
    public async Task GetSlideXmlAsync_ReturnsXmlString()
    {
        var path = CreateTempPptx();
        var xml = await _service.GetSlideXmlAsync(path, 0);
        Assert.NotNull(xml);
        Assert.Contains("sld", xml);
    }

    [Fact]
    public async Task GetSlideXmlAsync_OutOfRange_ThrowsException()
    {
        var path = CreateTempPptx();
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
            _service.GetSlideXmlAsync(path, 99));
    }

    [Fact]
    public async Task GetSlideContentAsync_ReturnsSlideIndex()
    {
        var path = CreateTempPptx();
        var content = await _service.GetSlideContentAsync(path, 0);
        Assert.Equal(0, content.SlideIndex);
    }

    [Fact]
    public async Task GetSlideContentAsync_ReturnsSlideDimensions()
    {
        var path = CreateTempPptx();
        var content = await _service.GetSlideContentAsync(path, 0);
        Assert.True(content.SlideWidthEmu > 0);
        Assert.True(content.SlideHeightEmu > 0);
    }

    [Fact]
    public async Task GetSlideContentAsync_ReturnsShapes()
    {
        var path = CreateTempPptx();
        var content = await _service.GetSlideContentAsync(path, 0);
        Assert.NotEmpty(content.Shapes);
    }

    [Fact]
    public async Task GetSlideContentAsync_TitleShapeHasText()
    {
        var path = CreateTempPptx("My Title");
        var content = await _service.GetSlideContentAsync(path, 0);
        var titleShape = content.Shapes.FirstOrDefault(s => s.IsPlaceholder);
        Assert.NotNull(titleShape);
        Assert.Equal("My Title", titleShape.Text);
    }

    [Fact]
    public async Task GetSlideContentAsync_TitleShapeIsTextType()
    {
        var path = CreateTempPptx();
        var content = await _service.GetSlideContentAsync(path, 0);
        var titleShape = content.Shapes.FirstOrDefault(s => s.IsPlaceholder);
        Assert.NotNull(titleShape);
        Assert.Equal("Text", titleShape.ShapeType);
    }

    [Fact]
    public async Task GetSlideContentAsync_TitleShapeHasPlaceholderType()
    {
        var path = CreateTempPptx();
        var content = await _service.GetSlideContentAsync(path, 0);
        var titleShape = content.Shapes.FirstOrDefault(s => s.IsPlaceholder);
        Assert.NotNull(titleShape);
        Assert.NotNull(titleShape.PlaceholderType);
    }

    [Fact]
    public async Task GetSlideContentAsync_OutOfRange_ThrowsException()
    {
        var path = CreateTempPptx();
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
            _service.GetSlideContentAsync(path, 99));
    }
}
