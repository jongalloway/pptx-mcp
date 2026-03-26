using System.Text.Json;

namespace PptxTools.Tests.Tools;

/// <summary>
/// Tool-level tests for pptx_validate_presentation (Issue #121 — Presentation validation and diagnostics).
/// Validates JSON structure, action routing, error handling, and parameter passthrough.
/// </summary>
[Trait("Category", "Integration")]
public class ValidationToolsTests : PptxTestBase
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public ValidationToolsTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    // ────────────────────────────────────────────────────────
    // Happy path: returns structured JSON
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Validate_ValidFile_ReturnsSuccessJson()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_validate_presentation(path, ValidationAction.Validate);

        var parsed = JsonSerializer.Deserialize<ValidationResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("Validate", parsed.Action);
    }

    [Fact]
    public async Task Validate_ValidFile_ReturnsZeroIssues()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_validate_presentation(path, ValidationAction.Validate);

        var parsed = JsonSerializer.Deserialize<ValidationResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.Equal(0, parsed.IssueCount);
        Assert.Empty(parsed.Issues);
    }

    [Fact]
    public async Task Validate_ValidFile_JsonIsIndented()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_validate_presentation(path, ValidationAction.Validate);

        // Indented JSON has newlines
        Assert.Contains(Environment.NewLine, result);
    }

    // ────────────────────────────────────────────────────────
    // File not found: structured error
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Validate_FileNotFound_ReturnsStructuredError()
    {
        var fakePath = @"C:\does-not-exist\file.pptx";

        var result = await _tools.pptx_validate_presentation(fakePath, ValidationAction.Validate);

        var parsed = JsonSerializer.Deserialize<ValidationResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("File not found", parsed.Message);
    }

    [Fact]
    public async Task Validate_FileNotFound_ReturnsZeroIssueCount()
    {
        var fakePath = @"C:\does-not-exist\file.pptx";

        var result = await _tools.pptx_validate_presentation(fakePath, ValidationAction.Validate);

        var parsed = JsonSerializer.Deserialize<ValidationResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.Equal(0, parsed.IssueCount);
        Assert.Equal(0, parsed.ErrorCount);
        Assert.Empty(parsed.Issues);
    }

    // ────────────────────────────────────────────────────────
    // Slide number filter passthrough
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Validate_WithSlideNumber_PassesThroughToService()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" });

        var resultAll = await _tools.pptx_validate_presentation(path, ValidationAction.Validate);
        var resultSlide1 = await _tools.pptx_validate_presentation(path, ValidationAction.Validate, slideNumber: 1);

        var parsedAll = JsonSerializer.Deserialize<ValidationResult>(resultAll, JsonOptions);
        var parsedSlide1 = JsonSerializer.Deserialize<ValidationResult>(resultSlide1, JsonOptions);

        Assert.NotNull(parsedAll);
        Assert.NotNull(parsedSlide1);

        // Filtered should have equal or fewer issues (no cross-slide duplicates)
        Assert.True(parsedSlide1.IssueCount <= parsedAll.IssueCount);
    }

    [Fact]
    public async Task Validate_WithSlideNumber_NullDefault_ValidatesAllSlides()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_validate_presentation(path, ValidationAction.Validate, slideNumber: null);

        var parsed = JsonSerializer.Deserialize<ValidationResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
    }

    // ────────────────────────────────────────────────────────
    // JSON structure: all expected fields present
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Validate_ResponseJson_HasAllExpectedFields()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_validate_presentation(path, ValidationAction.Validate);

        using var jsonDoc = JsonDocument.Parse(result);
        var root = jsonDoc.RootElement;

        Assert.True(root.TryGetProperty("Success", out _));
        Assert.True(root.TryGetProperty("Action", out _));
        Assert.True(root.TryGetProperty("IssueCount", out _));
        Assert.True(root.TryGetProperty("ErrorCount", out _));
        Assert.True(root.TryGetProperty("WarningCount", out _));
        Assert.True(root.TryGetProperty("InfoCount", out _));
        Assert.True(root.TryGetProperty("Issues", out _));
        Assert.True(root.TryGetProperty("Message", out _));
    }

    [Fact]
    public async Task Validate_WithIssues_IssueJsonHasAllFields()
    {
        var path = CreatePptxWithDuplicateShapeIds();

        var result = await _tools.pptx_validate_presentation(path, ValidationAction.Validate);

        using var jsonDoc = JsonDocument.Parse(result);
        var issues = jsonDoc.RootElement.GetProperty("Issues");
        Assert.True(issues.GetArrayLength() > 0);

        var firstIssue = issues[0];
        Assert.True(firstIssue.TryGetProperty("SlideNumber", out _));
        Assert.True(firstIssue.TryGetProperty("Severity", out _));
        Assert.True(firstIssue.TryGetProperty("Category", out _));
        Assert.True(firstIssue.TryGetProperty("Description", out _));
        Assert.True(firstIssue.TryGetProperty("Recommendation", out _));
    }

    // ────────────────────────────────────────────────────────
    // Validate action enum
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Validate_ActionEnum_ReturnsActionInResponse()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_validate_presentation(path, ValidationAction.Validate);

        using var jsonDoc = JsonDocument.Parse(result);
        var action = jsonDoc.RootElement.GetProperty("Action").GetString();
        Assert.Equal("Validate", action);
    }

    // ────────────────────────────────────────────────────────
    // Multi-slide with issues: tool returns all detected issues
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Validate_PresentationWithIssues_ReturnsNonZeroIssueCount()
    {
        var path = CreatePptxWithDuplicateShapeIds();

        var result = await _tools.pptx_validate_presentation(path, ValidationAction.Validate);

        var parsed = JsonSerializer.Deserialize<ValidationResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.True(parsed.IssueCount > 0);
        Assert.True(parsed.ErrorCount > 0);
    }

    // ════════════════════════════════════════════════════════
    // Fixture helpers
    // ════════════════════════════════════════════════════════

    private string CreatePptxWithDuplicateShapeIds()
    {
        var path = CreateMinimalPptx("Dup Test");

        using var doc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, true);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        uint duplicateId = 2;
        foreach (var child in shapeTree.Elements<DocumentFormat.OpenXml.Presentation.Shape>())
        {
            var id = child.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (id.HasValue) { duplicateId = id.Value; break; }
        }

        shapeTree.Append(new DocumentFormat.OpenXml.Presentation.Shape(
            new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = duplicateId, Name = "DupShape" },
                new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(),
                new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
            new DocumentFormat.OpenXml.Presentation.ShapeProperties(
                new DocumentFormat.OpenXml.Drawing.Transform2D(
                    new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
                    new DocumentFormat.OpenXml.Drawing.Extents { Cx = 914400, Cy = 457200 })),
            new DocumentFormat.OpenXml.Presentation.TextBody(
                new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                new DocumentFormat.OpenXml.Drawing.Paragraph(
                    new DocumentFormat.OpenXml.Drawing.Run(
                        new DocumentFormat.OpenXml.Drawing.Text("Dup"))))));

        slidePart.Slide.Save();
        return path;
    }
}
