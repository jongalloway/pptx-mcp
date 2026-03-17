using System.Text.Json;
using DocumentFormat.OpenXml.Presentation;
using PptxMcp.Models;

namespace PptxMcp.Tests.Tools;

public class PptxToolsTests : PptxTestBase
{
    private readonly PptxTools _tools;

    public PptxToolsTests()
    {
        _tools = new PptxTools(Service);
    }

    private string CreateTemplatePptx()
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);
        TemplateDeckHelper.CreateTemplatePresentation(path);
        return path;
    }

    [Fact]
    public async Task pptx_list_slides_ReturnsJson()
    {
        var path = CreateMinimalPptx();
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
        var path = CreateMinimalPptx();
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
        var path = CreateMinimalPptx();
        var result = await _tools.pptx_add_slide(path);
        Assert.Contains("successfully", result);
    }

    [Fact]
    public async Task pptx_update_text_ReturnsSuccessMessage()
    {
        var path = CreateMinimalPptx();
        var result = await _tools.pptx_update_text(path, 0, 0, "New Text");
        Assert.Contains("successfully", result);
    }

    [Fact]
    public async Task pptx_add_slide_from_layout_ReturnsStructuredJson()
    {
        var path = CreateTemplatePptx();

        var result = await _tools.pptx_add_slide_from_layout(path, TemplateDeckHelper.TitleBodyLayoutName, new Dictionary<string, string>
        {
            ["Title"] = "Agenda",
            ["Body:1"] = "Wins",
            ["Body:2"] = "Risks"
        });
        var addResult = JsonSerializer.Deserialize<AddSlideFromLayoutResult>(result);

        Assert.NotNull(addResult);
        Assert.True(addResult.Success);
        Assert.Equal(2, addResult.SlideNumber);
        Assert.Equal(3, addResult.PlaceholdersPopulated);
    }

    [Fact]
    public async Task pptx_add_slide_from_layout_ReturnsStructuredFailureJson()
    {
        var path = CreateTemplatePptx();

        var result = await _tools.pptx_add_slide_from_layout(path, "Missing Layout");
        var addResult = JsonSerializer.Deserialize<AddSlideFromLayoutResult>(result);

        Assert.NotNull(addResult);
        Assert.False(addResult.Success);
        Assert.Contains("Missing Layout", addResult.Message);
    }

    [Fact]
    public async Task pptx_duplicate_slide_ReturnsStructuredJson()
    {
        var path = CreateTemplatePptx();

        var result = await _tools.pptx_duplicate_slide(path, 1, new Dictionary<string, string>
        {
            ["Title"] = "Duplicated Review"
        });
        var duplicateResult = JsonSerializer.Deserialize<DuplicateSlideResult>(result);

        Assert.NotNull(duplicateResult);
        Assert.True(duplicateResult.Success);
        Assert.Equal(2, duplicateResult.NewSlideNumber);
        Assert.Equal(1, duplicateResult.OverridesApplied);
    }

    [Fact]
    public async Task pptx_update_slide_data_ReturnsStructuredJson()
    {
        var path = CreatePptxWithSlides(
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

        var slideContent = Service.GetSlideContent(path, 0);
        var updatedShape = Assert.Single(slideContent.Shapes, shape => shape.Name == "Revenue Value");
        Assert.Equal(string.Empty, updatedShape.Text);
    }

    [Fact]
    public async Task pptx_update_slide_data_ReturnsStructuredFailureJson()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_update_slide_data(path, 3, "Missing Shape", null, "Updated");
        var updateResult = JsonSerializer.Deserialize<SlideDataUpdateResult>(result);

        Assert.NotNull(updateResult);
        Assert.False(updateResult.Success);
        Assert.Contains("out of range", updateResult.Message);
    }

    [Fact]
    public async Task pptx_batch_update_ReturnsStructuredJson()
    {
        var path = CreatePptxWithSlides(
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
            },
            new TestSlideDefinition
            {
                TitleText = "Risks",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Risk Body",
                        PlaceholderType = PlaceholderValues.Body,
                        Paragraphs = ["Legacy blocker"]
                    }
                ]
            });

        var result = await _tools.pptx_batch_update(path,
        [
            new BatchUpdateMutation(1, "Revenue Value", "15%"),
            new BatchUpdateMutation(2, "Risk Body", "Mitigate EMEA churn")
        ]);
        var batchResult = JsonSerializer.Deserialize<BatchUpdateResult>(result);

        Assert.NotNull(batchResult);
        Assert.Equal(2, batchResult.TotalMutations);
        Assert.Equal(2, batchResult.SuccessCount);
        Assert.Equal(0, batchResult.FailureCount);
        Assert.All(batchResult.Results, mutationResult =>
        {
            Assert.True(mutationResult.Success);
            Assert.Equal("shapeName", mutationResult.MatchedBy);
            Assert.Null(mutationResult.Error);
        });

        var firstSlideContent = Service.GetSlideContent(path, 0);
        var secondSlideContent = Service.GetSlideContent(path, 1);
        Assert.Equal("15%", Assert.Single(firstSlideContent.Shapes, shape => shape.Name == "Revenue Value").Text);
        Assert.Equal("Mitigate EMEA churn", Assert.Single(secondSlideContent.Shapes, shape => shape.Name == "Risk Body").Text);
    }

    [Fact]
    public async Task pptx_batch_update_FileNotFound_ReturnsStructuredFailureJson()
    {
        var result = await _tools.pptx_batch_update("C:\\does-not-exist\\file.pptx",
        [
            new BatchUpdateMutation(1, "Revenue Value", "Updated")
        ]);
        var batchResult = JsonSerializer.Deserialize<BatchUpdateResult>(result);

        Assert.NotNull(batchResult);
        Assert.Equal(1, batchResult.TotalMutations);
        Assert.Equal(0, batchResult.SuccessCount);
        Assert.Equal(1, batchResult.FailureCount);

        var mutationResult = Assert.Single(batchResult.Results);
        Assert.False(mutationResult.Success);
        Assert.Null(mutationResult.MatchedBy);
        Assert.Contains("File not found", mutationResult.Error);
    }

    [Fact]
    public async Task pptx_get_slide_xml_ReturnsXml()
    {
        var path = CreateMinimalPptx();
        var result = await _tools.pptx_get_slide_xml(path, 0);
        Assert.Contains("sld", result);
    }

    [Fact]
    public async Task pptx_insert_image_FileNotFound_ReturnsError()
    {
        var path = CreateMinimalPptx();
        var result = await _tools.pptx_insert_image(path, 0, "C:\\does-not-exist\\image.png");
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_get_slide_content_ReturnsJson()
    {
        var path = CreateMinimalPptx();
        var result = await _tools.pptx_get_slide_content(path, 0);
        Assert.Contains("SlideIndex", result);
        Assert.Contains("Shapes", result);
    }

    [Fact]
    public async Task pptx_get_slide_content_ContainsTitleText()
    {
        var path = CreateMinimalPptx();
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
        var path = CreateMinimalPptx();
        var result = await _tools.pptx_get_slide_content(path, 0);
        Assert.Contains("Text", result);
    }

    [Fact]
    public async Task pptx_extract_talking_points_ReturnsStructuredJson()
    {
        var path = CreatePptxWithSlides(
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
        var path = CreatePptxWithSlides(
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
        var path = CreateMinimalPptx();
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
        var path = CreateMinimalPptx();
        var result = await _tools.pptx_write_notes(path, 0, "Citation: https://example.com");
        Assert.False(result.StartsWith("Error:", StringComparison.Ordinal));
        var slides = Service.GetSlides(path);
        Assert.Equal("Citation: https://example.com", slides[0].Notes);
    }

    [Fact]
    public async Task pptx_write_notes_Append_PreservesExistingNotes()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Slide",
            SpeakerNotesText = "Original"
        });
        var result = await _tools.pptx_write_notes(path, 0, "Appended", append: true);
        Assert.False(result.StartsWith("Error:", StringComparison.Ordinal));
        var slides = Service.GetSlides(path);
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
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide A" },
            new TestSlideDefinition { TitleText = "Slide B" },
            new TestSlideDefinition { TitleText = "Slide C" });

        var result = await _tools.pptx_move_slide(path, 1, 3);

        Assert.Contains("successfully", result);
        var slides = Service.GetSlides(path);
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
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide A" },
            new TestSlideDefinition { TitleText = "Slide B" });

        var result = await _tools.pptx_move_slide(path, 5, 1);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_delete_slide_ReturnsSuccessMessage()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Keep" },
            new TestSlideDefinition { TitleText = "Delete Me" });

        var result = await _tools.pptx_delete_slide(path, 2);

        Assert.Contains("successfully", result);
        var slides = Service.GetSlides(path);
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
        var path = CreateMinimalPptx();
        var result = await _tools.pptx_delete_slide(path, 1);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_reorder_slides_ReturnsSuccessMessage()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "First" },
            new TestSlideDefinition { TitleText = "Second" },
            new TestSlideDefinition { TitleText = "Third" });

        var result = await _tools.pptx_reorder_slides(path, [3, 1, 2]);

        Assert.Contains("successfully", result);
        var slides = Service.GetSlides(path);
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
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });

        var result = await _tools.pptx_reorder_slides(path, [1, 1]);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_reorder_slides_WrongLength_ReturnsError()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" },
            new TestSlideDefinition { TitleText = "C" });

        var result = await _tools.pptx_reorder_slides(path, [1, 2]);
        Assert.StartsWith("Error:", result);
    }

    #region pptx_replace_image tool tests

    private static readonly byte[] ToolTestPngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nZxQAAAAASUVORK5CYII=");

    private string CreatePptxWithPicture()
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);
        TestPptxHelper.CreatePresentation(path, [new TestSlideDefinition { TitleText = "Slide With Image", IncludeImage = true }]);
        return path;
    }

    private string CreateTempImageFile(string extension = ".png")
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + extension);
        TrackTempFile(path);
        File.WriteAllBytes(path, ToolTestPngBytes);
        return path;
    }

    [Fact]
    public async Task pptx_replace_image_ReturnsStructuredJson()
    {
        var pptxPath = CreatePptxWithPicture();
        var imagePath = CreateTempImageFile();

        var result = await _tools.pptx_replace_image(pptxPath, 1, shapeIndex: 0, imagePath: imagePath);
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);

        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
    }

    [Fact]
    public async Task pptx_replace_image_FileNotFound_ReturnsStructuredError()
    {
        var imagePath = CreateTempImageFile();
        var result = await _tools.pptx_replace_image("C:\\does-not-exist\\file.pptx", 1, shapeIndex: 0, imagePath: imagePath);
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);

        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("File not found", parsed.Message);
    }

    [Fact]
    public async Task pptx_replace_image_ImageNotFound_ReturnsStructuredError()
    {
        var pptxPath = CreatePptxWithPicture();
        var result = await _tools.pptx_replace_image(pptxPath, 1, shapeIndex: 0, imagePath: "C:\\does-not-exist\\image.png");
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);

        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("Image file not found", parsed.Message);
    }

    [Fact]
    public async Task pptx_replace_image_WithAltText_ReturnsInResult()
    {
        var pptxPath = CreatePptxWithPicture();
        var imagePath = CreateTempImageFile();

        var result = await _tools.pptx_replace_image(pptxPath, 1, shapeIndex: 0, imagePath: imagePath, altText: "Team photo");
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);

        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("Team photo", parsed.AltText);
    }

    [Fact]
    public async Task pptx_replace_image_Exception_ReturnsStructuredError()
    {
        var pptxPath = CreatePptxWithPicture();
        var imagePath = CreateTempImageFile(".tiff");

        var result = await _tools.pptx_replace_image(pptxPath, 1, shapeIndex: 0, imagePath: imagePath);
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);

        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("Unsupported", parsed.Message);
    }

    #endregion
}
