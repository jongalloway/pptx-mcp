using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;

namespace PptxTools.Tests.Tools;

/// <summary>
/// Tool-level tests for pptx_export_json MCP tool (Issue #128 — Export presentation to structured JSON).
/// Validates JSON output structure, action routing, error handling, and structured error responses.
/// </summary>
[Trait("Category", "Integration")]
public class ExportJsonToolsTests : PptxTestBase
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public ExportJsonToolsTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    // ────────────────────────────────────────────────────────
    // Fixture helpers
    // ────────────────────────────────────────────────────────

    private string CreatePptxWithMetadata(string? title = null, string? creator = null)
    {
        var path = CreateMinimalPptx();

        using var doc = PresentationDocument.Open(path, true);
        if (title is not null) doc.PackageProperties.Title = title;
        if (creator is not null) doc.PackageProperties.Creator = creator;

        return path;
    }

    // ────────────────────────────────────────────────────────
    // Action routing: Full
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task ExportJson_Full_ReturnsStructuredResult()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_export_json(path, ExportJsonAction.Full);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("Full", parsed.Action);
    }

    [Fact]
    public async Task ExportJson_Full_ContainsSlidesAndMetadata()
    {
        var path = CreatePptxWithMetadata(title: "Test Deck");

        var result = await _tools.pptx_export_json(path, ExportJsonAction.Full);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.NotNull(parsed.Metadata);
        Assert.NotNull(parsed.Slides);
        Assert.True(parsed.Slides.Count > 0);
    }

    [Fact]
    public async Task ExportJson_Full_SlideCountMatches()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" });

        var result = await _tools.pptx_export_json(path, ExportJsonAction.Full);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.Equal(2, parsed.SlideCount);
    }

    // ────────────────────────────────────────────────────────
    // Action routing: SlidesOnly
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task ExportJson_SlidesOnly_ReturnsCorrectAction()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_export_json(path, ExportJsonAction.SlidesOnly);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.Equal("SlidesOnly", parsed.Action);
    }

    [Fact]
    public async Task ExportJson_SlidesOnly_SlidesPopulated()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "One" },
            new TestSlideDefinition { TitleText = "Two" },
            new TestSlideDefinition { TitleText = "Three" });

        var result = await _tools.pptx_export_json(path, ExportJsonAction.SlidesOnly);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.NotNull(parsed.Slides);
        Assert.Equal(3, parsed.Slides.Count);
    }

    // ────────────────────────────────────────────────────────
    // Action routing: MetadataOnly
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task ExportJson_MetadataOnly_ReturnsCorrectAction()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_export_json(path, ExportJsonAction.MetadataOnly);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.Equal("MetadataOnly", parsed.Action);
    }

    [Fact]
    public async Task ExportJson_MetadataOnly_MetadataPopulated()
    {
        var path = CreatePptxWithMetadata(title: "My Deck", creator: "Author");

        var result = await _tools.pptx_export_json(path, ExportJsonAction.MetadataOnly);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.NotNull(parsed.Metadata);
        Assert.Equal("My Deck", parsed.Metadata.Title);
        Assert.Equal("Author", parsed.Metadata.Creator);
    }

    // ────────────────────────────────────────────────────────
    // Action routing: SchemaOnly
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task ExportJson_SchemaOnly_ReturnsCorrectAction()
    {
        var result = await _tools.pptx_export_json(null, ExportJsonAction.SchemaOnly);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.Equal("SchemaOnly", parsed.Action);
    }

    [Fact]
    public async Task ExportJson_SchemaOnly_SchemaPopulated()
    {
        var result = await _tools.pptx_export_json(null, ExportJsonAction.SchemaOnly);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.NotNull(parsed.Schema);
        Assert.Contains("PresentationExport", parsed.Schema);
    }

    [Fact]
    public async Task ExportJson_SchemaOnly_WithFilePath_StillWorks()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_export_json(path, ExportJsonAction.SchemaOnly);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("SchemaOnly", parsed.Action);
    }

    // ────────────────────────────────────────────────────────
    // Error handling: file not found
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task ExportJson_FileNotFound_ReturnsStructuredError()
    {
        var fakePath = @"C:\does-not-exist\presentation.pptx";

        var result = await _tools.pptx_export_json(fakePath, ExportJsonAction.Full);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("not found", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(ExportJsonAction.Full)]
    [InlineData(ExportJsonAction.SlidesOnly)]
    [InlineData(ExportJsonAction.MetadataOnly)]
    public async Task ExportJson_FileNotFound_AllNonSchemaActions_ReturnsStructuredError(ExportJsonAction action)
    {
        var fakePath = @"C:\does-not-exist\missing.pptx";

        var result = await _tools.pptx_export_json(fakePath, action);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task ExportJson_NullFilePath_NonSchemaAction_ReturnsError()
    {
        var result = await _tools.pptx_export_json(null, ExportJsonAction.Full);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("filePath", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ExportJson_EmptyFilePath_NonSchemaAction_ReturnsError()
    {
        var result = await _tools.pptx_export_json("", ExportJsonAction.Full);

        var parsed = JsonSerializer.Deserialize<PresentationExport>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    // ────────────────────────────────────────────────────────
    // JSON structure validation
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task ExportJson_ResponseJson_HasAllExpectedFields()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_export_json(path, ExportJsonAction.Full);

        using var jsonDoc = JsonDocument.Parse(result);
        var root = jsonDoc.RootElement;

        Assert.True(root.TryGetProperty("Success", out _));
        Assert.True(root.TryGetProperty("Action", out _));
        Assert.True(root.TryGetProperty("FilePath", out _));
        Assert.True(root.TryGetProperty("SlideCount", out _));
        Assert.True(root.TryGetProperty("Message", out _));
        Assert.True(root.TryGetProperty("Metadata", out _));
        Assert.True(root.TryGetProperty("Slides", out _));
        Assert.True(root.TryGetProperty("Schema", out _));
    }

    [Fact]
    public async Task ExportJson_ResponseJson_IsIndented()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_export_json(path, ExportJsonAction.Full);

        Assert.Contains(Environment.NewLine, result);
    }

    [Fact]
    public async Task ExportJson_ErrorJson_HasExpectedFields()
    {
        var fakePath = @"C:\does-not-exist\error.pptx";

        var result = await _tools.pptx_export_json(fakePath, ExportJsonAction.Full);

        using var jsonDoc = JsonDocument.Parse(result);
        var root = jsonDoc.RootElement;

        Assert.True(root.TryGetProperty("Success", out _));
        Assert.True(root.TryGetProperty("Action", out _));
        Assert.True(root.TryGetProperty("Message", out _));
    }

    [Fact]
    public async Task ExportJson_Full_SlideJsonHasExpectedFields()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Test Slide" });

        var result = await _tools.pptx_export_json(path, ExportJsonAction.Full);

        using var jsonDoc = JsonDocument.Parse(result);
        var slides = jsonDoc.RootElement.GetProperty("Slides");
        Assert.True(slides.GetArrayLength() > 0);
        var firstSlide = slides[0];
        Assert.True(firstSlide.TryGetProperty("SlideNumber", out _));
        Assert.True(firstSlide.TryGetProperty("SlideIndex", out _));
        Assert.True(firstSlide.TryGetProperty("Shapes", out _));
        Assert.True(firstSlide.TryGetProperty("SpeakerNotes", out _));
    }

    [Fact]
    public async Task ExportJson_MetadataOnly_MetadataJsonHasExpectedFields()
    {
        var path = CreatePptxWithMetadata(title: "Doc Title", creator: "Author");

        var result = await _tools.pptx_export_json(path, ExportJsonAction.MetadataOnly);

        using var jsonDoc = JsonDocument.Parse(result);
        var metadata = jsonDoc.RootElement.GetProperty("Metadata");
        Assert.True(metadata.TryGetProperty("Title", out _));
        Assert.True(metadata.TryGetProperty("Creator", out _));
        Assert.True(metadata.TryGetProperty("Subject", out _));
        Assert.True(metadata.TryGetProperty("Keywords", out _));
        Assert.True(metadata.TryGetProperty("Category", out _));
    }

    [Fact]
    public async Task ExportJson_WithTable_ShapeJsonHasTableField()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                Tables =
                [
                    new TestTableDefinition
                    {
                        Name = "Test Table",
                        Rows = [["A", "B"]]
                    }
                ]
            });

        var result = await _tools.pptx_export_json(path, ExportJsonAction.Full);

        using var jsonDoc = JsonDocument.Parse(result);
        var slides = jsonDoc.RootElement.GetProperty("Slides");
        Assert.True(slides.GetArrayLength() > 0);
        var shapes = slides[0].GetProperty("Shapes");
        bool hasTable = false;
        for (int i = 0; i < shapes.GetArrayLength(); i++)
        {
            if (shapes[i].TryGetProperty("Table", out var table) &&
                table.ValueKind != JsonValueKind.Null)
            {
                hasTable = true;
                Assert.True(table.TryGetProperty("RowCount", out _));
                Assert.True(table.TryGetProperty("ColumnCount", out _));
                Assert.True(table.TryGetProperty("Cells", out _));
                break;
            }
        }
        Assert.True(hasTable, "Expected at least one shape with a Table property");
    }

    [Fact]
    public async Task ExportJson_WithImage_ShapeJsonHasImageField()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Image Slide",
                IncludeImage = true
            });

        var result = await _tools.pptx_export_json(path, ExportJsonAction.Full);

        using var jsonDoc = JsonDocument.Parse(result);
        var slides = jsonDoc.RootElement.GetProperty("Slides");
        Assert.True(slides.GetArrayLength() > 0);
        var shapes = slides[0].GetProperty("Shapes");
        bool hasImage = false;
        for (int i = 0; i < shapes.GetArrayLength(); i++)
        {
            if (shapes[i].TryGetProperty("Image", out var image) &&
                image.ValueKind != JsonValueKind.Null)
            {
                hasImage = true;
                Assert.True(image.TryGetProperty("ContentType", out _));
                Assert.True(image.TryGetProperty("RelationshipId", out _));
                break;
            }
        }
        Assert.True(hasImage, "Expected at least one shape with an Image property");
    }

    [Fact]
    public async Task ExportJson_WithChart_ShapeJsonHasChartField()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                Charts =
                [
                    new TestChartDefinition
                    {
                        Name = "Revenue",
                        ChartType = "Column",
                        Categories = ["Q1", "Q2"],
                        Series = [new TestSeriesDefinition { Name = "Rev", Values = [10, 20] }]
                    }
                ]
            });

        var result = await _tools.pptx_export_json(path, ExportJsonAction.Full);

        using var jsonDoc = JsonDocument.Parse(result);
        var slides = jsonDoc.RootElement.GetProperty("Slides");
        Assert.True(slides.GetArrayLength() > 0);
        var shapes = slides[0].GetProperty("Shapes");
        bool hasChart = false;
        for (int i = 0; i < shapes.GetArrayLength(); i++)
        {
            if (shapes[i].TryGetProperty("Chart", out var chart) &&
                chart.ValueKind != JsonValueKind.Null)
            {
                hasChart = true;
                Assert.True(chart.TryGetProperty("ChartType", out _));
                Assert.True(chart.TryGetProperty("SeriesCount", out _));
                break;
            }
        }
        Assert.True(hasChart, "Expected at least one shape with a Chart property");
    }
}
