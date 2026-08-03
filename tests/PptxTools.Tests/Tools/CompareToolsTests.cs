using System.Text.Json;

namespace PptxTools.Tests.Tools;

/// <summary>
/// Tool-level tests for pptx_compare_presentations MCP tool (Issue #133 — Compare two presentations).
/// Validates JSON output structure, action routing, dual-file error handling, and structured error responses.
/// </summary>
[Trait("Category", "Integration")]
public class CompareToolsTests : PptxTestBase
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public CompareToolsTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    // ────────────────────────────────────────────────────────
    // Fixture helpers
    // ────────────────────────────────────────────────────────

    private string CreatePptxWithMetadata(string? title = null, string? creator = null)
    {
        var path = CreateMinimalPptx();

        using var doc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, true);
        if (title is not null) doc.PackageProperties.Title = title;
        if (creator is not null) doc.PackageProperties.Creator = creator;

        return path;
    }

    // ────────────────────────────────────────────────────────
    // Action routing: Full
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Compare_Full_ReturnsStructuredResult()
    {
        var source = CreateMinimalPptx();
        var target = CreateMinimalPptx();

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.Full);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("Full", parsed.Action);
    }

    [Fact]
    public async Task Compare_Full_IdenticalFiles_ReturnsAreIdenticalTrue()
    {
        var source = CreateMinimalPptx();
        var target = CreateMinimalPptx();

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.Full);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.True(parsed.AreIdentical);
        Assert.Equal(0, parsed.DifferenceCount);
    }

    [Fact]
    public async Task Compare_Full_DifferentFiles_ReturnsAreIdenticalFalse()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Version 1" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Version 2" });

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.Full);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.AreIdentical);
        Assert.True(parsed.DifferenceCount > 0);
    }

    // ────────────────────────────────────────────────────────
    // Action routing: SlidesOnly
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Compare_SlidesOnly_ReturnsCorrectAction()
    {
        var source = CreateMinimalPptx();
        var target = CreateMinimalPptx();

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.SlidesOnly);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.Equal("SlidesOnly", parsed.Action);
    }

    [Fact]
    public async Task Compare_SlidesOnly_DifferentSlideCount_ReportsDifferences()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" });

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.SlidesOnly);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.AreIdentical);
    }

    // ────────────────────────────────────────────────────────
    // Action routing: TextOnly
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Compare_TextOnly_ReturnsCorrectAction()
    {
        var source = CreateMinimalPptx();
        var target = CreateMinimalPptx();

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.TextOnly);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.Equal("TextOnly", parsed.Action);
    }

    [Fact]
    public async Task Compare_TextOnly_TextChanged_ReportsDifferences()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Original" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Modified" });

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.TextOnly);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.AreIdentical);
    }

    // ────────────────────────────────────────────────────────
    // Action routing: MetadataOnly
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Compare_MetadataOnly_ReturnsCorrectAction()
    {
        var source = CreateMinimalPptx();
        var target = CreateMinimalPptx();

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.MetadataOnly);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.Equal("MetadataOnly", parsed.Action);
    }

    [Fact]
    public async Task Compare_MetadataOnly_DifferentTitle_ReportsDifferences()
    {
        var source = CreatePptxWithMetadata(title: "Deck A");
        var target = CreatePptxWithMetadata(title: "Deck B");

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.MetadataOnly);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.AreIdentical);
    }

    // ────────────────────────────────────────────────────────
    // Error handling: file not found
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Compare_SourceFileNotFound_ReturnsStructuredError()
    {
        var fakePath = @"C:\does-not-exist\source.pptx";
        var target = CreateMinimalPptx();

        var result = await _tools.pptx_compare_presentations(fakePath, target, CompareAction.Full);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("not found", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Compare_TargetFileNotFound_ReturnsStructuredError()
    {
        var source = CreateMinimalPptx();
        var fakePath = @"C:\does-not-exist\target.pptx";

        var result = await _tools.pptx_compare_presentations(source, fakePath, CompareAction.Full);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("not found", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Compare_BothFilesNotFound_ReturnsStructuredError()
    {
        var fakeSource = @"C:\does-not-exist\source.pptx";
        var fakeTarget = @"C:\does-not-exist\target.pptx";

        var result = await _tools.pptx_compare_presentations(fakeSource, fakeTarget, CompareAction.Full);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("not found", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(CompareAction.Full)]
    [InlineData(CompareAction.SlidesOnly)]
    [InlineData(CompareAction.TextOnly)]
    [InlineData(CompareAction.MetadataOnly)]
    public async Task Compare_SourceFileNotFound_AllActions_ReturnsStructuredError(CompareAction action)
    {
        var fakePath = @"C:\does-not-exist\source.pptx";
        var target = CreateMinimalPptx();

        var result = await _tools.pptx_compare_presentations(fakePath, target, action);

        var parsed = JsonSerializer.Deserialize<ComparisonResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    // ────────────────────────────────────────────────────────
    // JSON structure validation
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task Compare_ResponseJson_HasAllExpectedFields()
    {
        var source = CreateMinimalPptx();
        var target = CreateMinimalPptx();

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.Full);

        using var jsonDoc = JsonDocument.Parse(result);
        var root = jsonDoc.RootElement;

        Assert.True(root.TryGetProperty("Success", out _));
        Assert.True(root.TryGetProperty("Action", out _));
        Assert.True(root.TryGetProperty("SourceFile", out _));
        Assert.True(root.TryGetProperty("TargetFile", out _));
        Assert.True(root.TryGetProperty("AreIdentical", out _));
        Assert.True(root.TryGetProperty("DifferenceCount", out _));
        Assert.True(root.TryGetProperty("SlideDifferences", out _));
        Assert.True(root.TryGetProperty("TextDifferences", out _));
        Assert.True(root.TryGetProperty("MetadataDifferences", out _));
        Assert.True(root.TryGetProperty("Message", out _));
    }

    [Fact]
    public async Task Compare_ResponseJson_IsIndented()
    {
        var source = CreateMinimalPptx();
        var target = CreateMinimalPptx();

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.Full);

        Assert.Contains(Environment.NewLine, result);
    }

    [Fact]
    public async Task Compare_ErrorJson_HasAllExpectedFields()
    {
        var fakePath = @"C:\does-not-exist\source.pptx";
        var target = CreateMinimalPptx();

        var result = await _tools.pptx_compare_presentations(fakePath, target, CompareAction.Full);

        using var jsonDoc = JsonDocument.Parse(result);
        var root = jsonDoc.RootElement;

        Assert.True(root.TryGetProperty("Success", out _));
        Assert.True(root.TryGetProperty("Action", out _));
        Assert.True(root.TryGetProperty("Message", out _));
        Assert.True(root.TryGetProperty("DifferenceCount", out _));
    }

    [Fact]
    public async Task Compare_WithDifferences_SlideDifferencesJsonHasFields()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" });

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.Full);

        using var jsonDoc = JsonDocument.Parse(result);
        var slideDiffs = jsonDoc.RootElement.GetProperty("SlideDifferences");
        if (slideDiffs.GetArrayLength() > 0)
        {
            var firstDiff = slideDiffs[0];
            Assert.True(firstDiff.TryGetProperty("SlideNumber", out _));
            Assert.True(firstDiff.TryGetProperty("DifferenceType", out _));
        }
    }

    [Fact]
    public async Task Compare_WithTextDifferences_TextDifferencesJsonHasFields()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Alpha" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Beta" });

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.Full);

        using var jsonDoc = JsonDocument.Parse(result);
        var textDiffs = jsonDoc.RootElement.GetProperty("TextDifferences");
        if (textDiffs.GetArrayLength() > 0)
        {
            var firstDiff = textDiffs[0];
            Assert.True(firstDiff.TryGetProperty("SlideNumber", out _));
            Assert.True(firstDiff.TryGetProperty("ShapeName", out _));
        }
    }

    [Fact]
    public async Task Compare_WithMetadataDifferences_MetadataDifferencesJsonHasFields()
    {
        var source = CreatePptxWithMetadata(title: "A");
        var target = CreatePptxWithMetadata(title: "B");

        var result = await _tools.pptx_compare_presentations(source, target, CompareAction.MetadataOnly);

        using var jsonDoc = JsonDocument.Parse(result);
        var metaDiffs = jsonDoc.RootElement.GetProperty("MetadataDifferences");
        if (metaDiffs.GetArrayLength() > 0)
        {
            var firstDiff = metaDiffs[0];
            Assert.True(firstDiff.TryGetProperty("Property", out _));
            Assert.True(firstDiff.TryGetProperty("SourceValue", out _));
            Assert.True(firstDiff.TryGetProperty("TargetValue", out _));
        }
    }
}
