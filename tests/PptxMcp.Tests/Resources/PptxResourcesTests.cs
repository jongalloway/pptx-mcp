using System.Text.Json;
using ModelContextProtocol.Protocol;
using PptxMcp.Resources;

namespace PptxMcp.Tests.Resources;

public class PptxResourcesTests : PptxTestBase
{
    private readonly PptxResources _resources;

    public PptxResourcesTests()
    {
        _resources = new PptxResources(Service);
    }

    // --- GetSlides resource ---

    [Fact]
    public void GetSlides_ValidFile_ReturnsTextResourceContents()
    {
        var path = CreateMinimalPptx("My Title");
        var result = _resources.GetSlides(path);

        Assert.NotNull(result);
        Assert.Equal("application/json", result.MimeType);
        Assert.Contains("My Title", result.Text);
    }

    [Fact]
    public void GetSlides_ValidFile_ReturnsJsonArray()
    {
        var path = CreateMinimalPptx();
        var result = _resources.GetSlides(path);

        var doc = JsonDocument.Parse(result.Text!);
        Assert.Equal(JsonValueKind.Array, doc.RootElement.ValueKind);
        Assert.Equal(1, doc.RootElement.GetArrayLength());
    }

    [Fact]
    public void GetSlides_ValidFile_SlideHasExpectedFields()
    {
        var path = CreateMinimalPptx("Slide One");
        var result = _resources.GetSlides(path);

        var doc = JsonDocument.Parse(result.Text!);
        var slide = doc.RootElement[0];
        Assert.True(slide.TryGetProperty("Index", out _));
        Assert.True(slide.TryGetProperty("Title", out var titleProp));
        Assert.Equal("Slide One", titleProp.GetString());
    }

    [Fact]
    public void GetSlides_FileNotFound_ReturnsErrorJson()
    {
        var result = _resources.GetSlides("/nonexistent/path/deck.pptx");

        var doc = JsonDocument.Parse(result.Text!);
        Assert.True(doc.RootElement.TryGetProperty("error", out var errorProp));
        Assert.Contains("not found", errorProp.GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetSlides_UrlEncodedPath_DecodesCorrectly()
    {
        var path = CreateMinimalPptx("Encoded Title");
        var encoded = Uri.EscapeDataString(path);
        var result = _resources.GetSlides(encoded);

        var doc = JsonDocument.Parse(result.Text!);
        Assert.Equal(JsonValueKind.Array, doc.RootElement.ValueKind);
        Assert.Contains("Encoded Title", result.Text);
    }

    [Fact]
    public void GetSlides_UriIncludesEncodedFile()
    {
        var path = CreateMinimalPptx();
        var encoded = Uri.EscapeDataString(path);
        var result = _resources.GetSlides(encoded);

        Assert.Contains(encoded, result.Uri);
        Assert.EndsWith("/slides", result.Uri);
    }

    // --- GetLayouts resource ---

    [Fact]
    public void GetLayouts_ValidFile_ReturnsTextResourceContents()
    {
        var path = CreateMinimalPptx();
        var result = _resources.GetLayouts(path);

        Assert.NotNull(result);
        Assert.Equal("application/json", result.MimeType);
        Assert.NotNull(result.Text);
    }

    [Fact]
    public void GetLayouts_ValidFile_ReturnsJsonArrayWithNames()
    {
        var path = CreateMinimalPptx();
        var result = _resources.GetLayouts(path);

        var doc = JsonDocument.Parse(result.Text!);
        Assert.Equal(JsonValueKind.Array, doc.RootElement.ValueKind);
        var first = doc.RootElement[0];
        Assert.True(first.TryGetProperty("Name", out _));
        Assert.True(first.TryGetProperty("Index", out _));
    }

    [Fact]
    public void GetLayouts_FileNotFound_ReturnsErrorJson()
    {
        var result = _resources.GetLayouts("/nonexistent/path/deck.pptx");

        var doc = JsonDocument.Parse(result.Text!);
        Assert.True(doc.RootElement.TryGetProperty("error", out _));
    }

    [Fact]
    public void GetLayouts_UriEndsWithLayouts()
    {
        var path = CreateMinimalPptx();
        var encoded = Uri.EscapeDataString(path);
        var result = _resources.GetLayouts(encoded);

        Assert.Contains(encoded, result.Uri);
        Assert.EndsWith("/layouts", result.Uri);
    }

    // --- GetShapeMap resource ---

    [Fact]
    public void GetShapeMap_ValidFile_ReturnsTextResourceContents()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Slide 1",
                TextShapes =
                [
                    new TestTextShapeDefinition { Name = "Revenue Value", Paragraphs = ["$1M"] }
                ]
            });

        var result = _resources.GetShapeMap(path);

        Assert.NotNull(result);
        Assert.Equal("application/json", result.MimeType);
        Assert.NotNull(result.Text);
    }

    [Fact]
    public void GetShapeMap_ValidFile_ContainsSlideKeys()
    {
        var path = CreateMinimalPptx("Title Slide");
        var result = _resources.GetShapeMap(path);

        var doc = JsonDocument.Parse(result.Text!);
        Assert.Equal(JsonValueKind.Object, doc.RootElement.ValueKind);
        Assert.True(doc.RootElement.TryGetProperty("0", out _));
    }

    [Fact]
    public void GetShapeMap_ValidFile_ShapeHasExpectedFields()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Title",
                TextShapes =
                [
                    new TestTextShapeDefinition { Name = "KPI Shape", Paragraphs = ["100%"] }
                ]
            });

        var result = _resources.GetShapeMap(path);
        var doc = JsonDocument.Parse(result.Text!);
        var shapes = doc.RootElement.GetProperty("0");

        var kpiShape = Assert.Single(shapes.EnumerateArray(),
            shape => shape.TryGetProperty("Name", out var nameProp) && nameProp.GetString() == "KPI Shape");
        Assert.True(kpiShape.TryGetProperty("ShapeType", out _));
        Assert.True(kpiShape.TryGetProperty("Text", out _));
    }

    [Fact]
    public void GetShapeMap_FileNotFound_ReturnsErrorJson()
    {
        var result = _resources.GetShapeMap("/nonexistent/path/deck.pptx");

        var doc = JsonDocument.Parse(result.Text!);
        Assert.True(doc.RootElement.TryGetProperty("error", out _));
    }

    [Fact]
    public void GetShapeMap_UriEndsWithShapeMap()
    {
        var path = CreateMinimalPptx();
        var encoded = Uri.EscapeDataString(path);
        var result = _resources.GetShapeMap(encoded);

        Assert.Contains(encoded, result.Uri);
        Assert.EndsWith("/shape-map", result.Uri);
    }
}
