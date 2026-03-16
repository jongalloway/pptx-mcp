using System.Text.Json;
using DocumentFormat.OpenXml.Presentation;
using PptxMcp.Models;

namespace PptxMcp.Tests.Tools;

public class PptxToolsTests : IDisposable
{
    private readonly PresentationService _service = new();
    private readonly PptxTools _tools;
    private readonly List<string> _tempFiles = new();

    public PptxToolsTests()
    {
        _tools = new PptxTools(_service);
    }

    public void Dispose()
    {
        foreach (var file in _tempFiles)
            if (File.Exists(file)) File.Delete(file);
    }

    private string CreateTempPptx()
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        _tempFiles.Add(path);
        TestPptxHelper.CreateMinimalPresentation(path, "Test Slide");
        return path;
    }

    private string CreateCustomPptx(params TestSlideDefinition[] slides)
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        _tempFiles.Add(path);
        TestPptxHelper.CreatePresentation(path, slides);
        return path;
    }

    [Fact]
    public async Task pptx_list_slides_ReturnsJson()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_list_slides(path);
        Assert.Contains("Index", result);
    }

    [Fact]
    public async Task pptx_list_slides_FileNotFound_ReturnsError()
    {
        var result = await _tools.pptx_list_slides("C:\\does-not-exist\\file.pptx");
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_list_layouts_ReturnsJson()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_list_layouts(path);
        Assert.Contains("Name", result);
    }

    [Fact]
    public async Task pptx_list_layouts_FileNotFound_ReturnsError()
    {
        var result = await _tools.pptx_list_layouts("C:\\does-not-exist\\file.pptx");
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_add_slide_ReturnsSuccessMessage()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_add_slide(path);
        Assert.Contains("successfully", result);
    }

    [Fact]
    public async Task pptx_update_text_ReturnsSuccessMessage()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_update_text(path, 0, 0, "New Text");
        Assert.Contains("successfully", result);
    }

    [Fact]
    public async Task pptx_update_slide_data_ReturnsStructuredJson()
    {
        var path = CreateCustomPptx(
            new TestSlideDefinition
            {
                TitleText = "Dashboard",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Revenue Value",
                        PlaceholderType = PlaceholderValues.Body,
                        Paragraphs = ["12%"]
                    }
                ]
            });

        var result = await _tools.pptx_update_slide_data(path, 1, "Revenue Value", null, string.Empty);
        var updateResult = JsonSerializer.Deserialize<SlideDataUpdateResult>(result);

        Assert.NotNull(updateResult);
        Assert.True(updateResult.Success);
        Assert.Equal("Revenue Value", updateResult.ResolvedShapeName);
        Assert.Equal(string.Empty, updateResult.NewText);

        var slideContent = _service.GetSlideContent(path, 0);
        var updatedShape = Assert.Single(slideContent.Shapes, shape => shape.Name == "Revenue Value");
        Assert.Equal(string.Empty, updatedShape.Text);
    }

    [Fact]
    public async Task pptx_update_slide_data_ReturnsStructuredFailureJson()
    {
        var path = CreateTempPptx();

        var result = await _tools.pptx_update_slide_data(path, 3, "Missing Shape", null, "Updated");
        var updateResult = JsonSerializer.Deserialize<SlideDataUpdateResult>(result);

        Assert.NotNull(updateResult);
        Assert.False(updateResult.Success);
        Assert.Contains("out of range", updateResult.Message);
    }

    [Fact]
    public async Task pptx_get_slide_xml_ReturnsXml()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_get_slide_xml(path, 0);
        Assert.Contains("sld", result);
    }

    [Fact]
    public async Task pptx_insert_image_FileNotFound_ReturnsError()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_insert_image(path, 0, "C:\\does-not-exist\\image.png");
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_get_slide_content_ReturnsJson()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_get_slide_content(path, 0);
        Assert.Contains("SlideIndex", result);
        Assert.Contains("Shapes", result);
    }

    [Fact]
    public async Task pptx_get_slide_content_ContainsTitleText()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_get_slide_content(path, 0);
        Assert.Contains("Test Slide", result);
    }

    [Fact]
    public async Task pptx_get_slide_content_FileNotFound_ReturnsError()
    {
        var result = await _tools.pptx_get_slide_content("C:\\does-not-exist\\file.pptx", 0);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_get_slide_content_ContainsShapeType()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_get_slide_content(path, 0);
        Assert.Contains("Text", result);
    }

    [Fact]
    public async Task pptx_extract_talking_points_ReturnsStructuredJson()
    {
        var path = CreateCustomPptx(
            new TestSlideDefinition
            {
                TitleText = "Launch Review",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Body 1",
                        PlaceholderType = PlaceholderValues.Body,
                        Paragraphs =
                        [
                            "Launch completed in 3 regions",
                            "Error rate dropped below 1%"
                        ]
                    }
                ]
            });

        var result = await _tools.pptx_extract_talking_points(path, 2);
        var talkingPoints = JsonSerializer.Deserialize<List<SlideTalkingPoints>>(result);

        Assert.NotNull(talkingPoints);
        Assert.Single(talkingPoints);
        Assert.Equal(0, talkingPoints[0].SlideIndex);
        Assert.Equal(
            [
                "Launch completed in 3 regions",
                "Error rate dropped below 1%"
            ],
            talkingPoints[0].Points);
    }

    [Fact]
    public async Task pptx_extract_talking_points_DefaultsToFivePointsPerSlide()
    {
        var path = CreateCustomPptx(
            new TestSlideDefinition
            {
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Body 1",
                        PlaceholderType = PlaceholderValues.Body,
                        Paragraphs =
                        [
                            "Point 1 has enough detail",
                            "Point 2 has enough detail",
                            "Point 3 has enough detail",
                            "Point 4 has enough detail",
                            "Point 5 has enough detail",
                            "Point 6 has enough detail"
                        ]
                    }
                ]
            });

        var result = await _tools.pptx_extract_talking_points(path);
        var talkingPoints = JsonSerializer.Deserialize<List<SlideTalkingPoints>>(result);

        Assert.NotNull(talkingPoints);
        Assert.Equal(5, talkingPoints[0].Points.Count);
        Assert.DoesNotContain("Point 6 has enough detail", talkingPoints[0].Points);
    }

    [Fact]
    public async Task pptx_extract_talking_points_InvalidTopN_ReturnsError()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_extract_talking_points(path, 0);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_extract_talking_points_FileNotFound_ReturnsError()
    {
        var result = await _tools.pptx_extract_talking_points("C:\\does-not-exist\\file.pptx");
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_write_notes_CreatesNotes()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_write_notes(path, 0, "Citation: https://example.com");
        Assert.False(result.StartsWith("Error:", StringComparison.Ordinal));
        var slides = _service.GetSlides(path);
        Assert.Equal("Citation: https://example.com", slides[0].Notes);
    }

    [Fact]
    public async Task pptx_write_notes_Append_PreservesExistingNotes()
    {
        var path = CreateCustomPptx(new TestSlideDefinition
        {
            TitleText = "Slide",
            SpeakerNotesText = "Original"
        });
        var result = await _tools.pptx_write_notes(path, 0, "Appended", append: true);
        Assert.False(result.StartsWith("Error:", StringComparison.Ordinal));
        var slides = _service.GetSlides(path);
        Assert.NotNull(slides[0].Notes);
        Assert.Contains("Original", slides[0].Notes);
        Assert.Contains("Appended", slides[0].Notes);
    }

    [Fact]
    public async Task pptx_write_notes_FileNotFound_ReturnsError()
    {
        var result = await _tools.pptx_write_notes("C:\\does-not-exist\\file.pptx", 0, "notes");
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_move_slide_ReturnsSuccessMessage()
    {
        var path = CreateCustomPptx(
            new TestSlideDefinition { TitleText = "Slide A" },
            new TestSlideDefinition { TitleText = "Slide B" },
            new TestSlideDefinition { TitleText = "Slide C" });

        var result = await _tools.pptx_move_slide(path, 1, 3);

        Assert.Contains("successfully", result);
        var slides = _service.GetSlides(path);
        Assert.Equal("Slide B", slides[0].Title);
        Assert.Equal("Slide C", slides[1].Title);
        Assert.Equal("Slide A", slides[2].Title);
    }

    [Fact]
    public async Task pptx_move_slide_FileNotFound_ReturnsError()
    {
        var result = await _tools.pptx_move_slide("C:\\does-not-exist\\file.pptx", 1, 2);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_move_slide_InvalidSlideNumber_ReturnsError()
    {
        var path = CreateCustomPptx(
            new TestSlideDefinition { TitleText = "Slide A" },
            new TestSlideDefinition { TitleText = "Slide B" });

        var result = await _tools.pptx_move_slide(path, 5, 1);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_delete_slide_ReturnsSuccessMessage()
    {
        var path = CreateCustomPptx(
            new TestSlideDefinition { TitleText = "Keep" },
            new TestSlideDefinition { TitleText = "Delete Me" });

        var result = await _tools.pptx_delete_slide(path, 2);

        Assert.Contains("successfully", result);
        var slides = _service.GetSlides(path);
        Assert.Single(slides);
        Assert.Equal("Keep", slides[0].Title);
    }

    [Fact]
    public async Task pptx_delete_slide_FileNotFound_ReturnsError()
    {
        var result = await _tools.pptx_delete_slide("C:\\does-not-exist\\file.pptx", 1);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_delete_slide_LastSlide_ReturnsError()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_delete_slide(path, 1);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_reorder_slides_ReturnsSuccessMessage()
    {
        var path = CreateCustomPptx(
            new TestSlideDefinition { TitleText = "First" },
            new TestSlideDefinition { TitleText = "Second" },
            new TestSlideDefinition { TitleText = "Third" });

        var result = await _tools.pptx_reorder_slides(path, [3, 1, 2]);

        Assert.Contains("successfully", result);
        var slides = _service.GetSlides(path);
        Assert.Equal("Third", slides[0].Title);
        Assert.Equal("First", slides[1].Title);
        Assert.Equal("Second", slides[2].Title);
    }

    [Fact]
    public async Task pptx_reorder_slides_FileNotFound_ReturnsError()
    {
        var result = await _tools.pptx_reorder_slides("C:\\does-not-exist\\file.pptx", [1, 2]);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_reorder_slides_InvalidOrder_ReturnsError()
    {
        var path = CreateCustomPptx(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });

        var result = await _tools.pptx_reorder_slides(path, [1, 1]);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_reorder_slides_WrongLength_ReturnsError()
    {
        var path = CreateCustomPptx(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" },
            new TestSlideDefinition { TitleText = "C" });

        var result = await _tools.pptx_reorder_slides(path, [1, 2]);
        Assert.StartsWith("Error:", result);
    }
}
