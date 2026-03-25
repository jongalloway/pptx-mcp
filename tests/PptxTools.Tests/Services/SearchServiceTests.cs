using DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests.Services;

[Trait("Category", "Unit")]
public class SearchServiceTests : PptxTestBase
{
    // ─── SearchText ──────────────────────────────────────────────────────────────

    [Fact]
    public void SearchText_EmptySearchText_ReturnsFailure()
    {
        var path = CreateMinimalPptx("Hello World");
        var result = Service.SearchText(path, "");
        Assert.False(result.Success);
        Assert.NotNull(result.Message);
    }

    [Fact]
    public void SearchText_NullSearchText_ReturnsFailure()
    {
        var path = CreateMinimalPptx("Hello World");
        var result = Service.SearchText(path, null!);
        Assert.False(result.Success);
        Assert.NotNull(result.Message);
    }

    [Fact]
    public void SearchText_MatchFound_ReturnsSuccess()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Hello World",
                TextShapes = [new TestTextShapeDefinition { Paragraphs = ["The quick brown fox"] }]
            });

        var result = Service.SearchText(path, "quick");

        Assert.True(result.Success);
        Assert.Single(result.Matches);
        Assert.Equal("quick", result.Matches[0].MatchedText);
    }

    [Fact]
    public void SearchText_NoMatch_ReturnsSuccessWithEmptyMatches()
    {
        var path = CreateMinimalPptx("Hello World");
        var result = Service.SearchText(path, "NotInPresentation");

        Assert.True(result.Success);
        Assert.Empty(result.Matches);
        Assert.Equal(0, result.TotalMatches);
    }

    [Fact]
    public void SearchText_CaseInsensitive_FindsUppercaseText()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes = [new TestTextShapeDefinition { Paragraphs = ["HELLO WORLD"] }]
            });

        var result = Service.SearchText(path, "hello world", caseSensitive: false);

        Assert.True(result.Success);
        Assert.Single(result.Matches);
    }

    [Fact]
    public void SearchText_CaseSensitive_DoesNotMatchDifferentCase()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes = [new TestTextShapeDefinition { Paragraphs = ["HELLO WORLD"] }]
            });

        var result = Service.SearchText(path, "hello world", caseSensitive: true);

        Assert.True(result.Success);
        Assert.Empty(result.Matches);
    }

    [Fact]
    public void SearchText_MultipleMatches_ReturnsAllMatches()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes = [new TestTextShapeDefinition { Paragraphs = ["cat and cat and cat"] }]
            });

        var result = Service.SearchText(path, "cat");

        Assert.True(result.Success);
        Assert.Equal(3, result.TotalMatches);
    }

    [Fact]
    public void SearchText_SlideNumber_SearchesOnlySpecifiedSlide()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide One" },
            new TestSlideDefinition { TitleText = "Slide Two" });

        var result = Service.SearchText(path, "Slide One", slideNumber: 1);

        Assert.True(result.Success);
        Assert.Equal(1, result.SlidesSearched);
        Assert.Single(result.Matches);
    }

    [Fact]
    public void SearchText_OutOfRangeSlideNumber_ReturnsFailure()
    {
        var path = CreateMinimalPptx("Test");
        var result = Service.SearchText(path, "Test", slideNumber: 99);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SearchText_SlideNumberZero_ReturnsFailure()
    {
        var path = CreateMinimalPptx("Test");
        var result = Service.SearchText(path, "Test", slideNumber: 0);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SearchText_ReturnsCorrectSlideNumber()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide One" },
            new TestSlideDefinition { TitleText = "Slide Two" });

        var result = Service.SearchText(path, "Slide Two");

        Assert.True(result.Success);
        Assert.Single(result.Matches);
        Assert.Equal(2, result.Matches[0].SlideNumber);
    }

    [Fact]
    public void SearchText_SearchesTableCells()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                Tables =
                [
                    new TestTableDefinition
                    {
                        Rows = [["Alpha", "Beta"], ["Gamma", "Delta"]]
                    }
                ]
            });

        var result = Service.SearchText(path, "Gamma");

        Assert.True(result.Success);
        Assert.Single(result.Matches);
    }

    // ─── SearchByRegex ────────────────────────────────────────────────────────────

    [Fact]
    public void SearchByRegex_EmptyPattern_ReturnsFailure()
    {
        var path = CreateMinimalPptx("Hello World");
        var result = Service.SearchByRegex(path, "");
        Assert.False(result.Success);
        Assert.NotNull(result.Message);
    }

    [Fact]
    public void SearchByRegex_InvalidPattern_ReturnsFailure()
    {
        var path = CreateMinimalPptx("Hello World");
        var result = Service.SearchByRegex(path, "[invalid(");
        Assert.False(result.Success);
        Assert.Contains("Invalid regex", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SearchByRegex_ValidPattern_ReturnsMatches()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes = [new TestTextShapeDefinition { Paragraphs = ["Revenue: $1,234"] }]
            });

        var result = Service.SearchByRegex(path, @"\$[\d,]+");

        Assert.True(result.Success);
        Assert.Single(result.Matches);
        Assert.Equal("$1,234", result.Matches[0].MatchedText);
    }

    [Fact]
    public void SearchByRegex_CaseInsensitive_MatchesAnyCase()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes = [new TestTextShapeDefinition { Paragraphs = ["HELLO world Hello"] }]
            });

        var result = Service.SearchByRegex(path, "hello", caseSensitive: false);

        Assert.True(result.Success);
        Assert.Equal(2, result.TotalMatches);
    }

    [Fact]
    public void SearchByRegex_CaseSensitive_OnlyMatchesExactCase()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes = [new TestTextShapeDefinition { Paragraphs = ["HELLO world Hello"] }]
            });

        var result = Service.SearchByRegex(path, "Hello", caseSensitive: true);

        Assert.True(result.Success);
        Assert.Single(result.Matches);
        Assert.Equal("Hello", result.Matches[0].MatchedText);
    }

    [Fact]
    public void SearchByRegex_OutOfRangeSlideNumber_ReturnsFailure()
    {
        var path = CreateMinimalPptx("Test");
        var result = Service.SearchByRegex(path, "test", slideNumber: 99);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SearchByRegex_SlideFilter_SearchesOnlySpecifiedSlide()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TextShapes = [new TestTextShapeDefinition { Paragraphs = ["target text"] }] },
            new TestSlideDefinition { TextShapes = [new TestTextShapeDefinition { Paragraphs = ["other text"] }] });

        var result = Service.SearchByRegex(path, "target", slideNumber: 1);

        Assert.True(result.Success);
        Assert.Equal(1, result.SlidesSearched);
        Assert.Single(result.Matches);
    }

    // ─── FindEmptyShapes ──────────────────────────────────────────────────────────

    [Fact]
    public void FindEmptyShapes_AllShapesFilled_ReturnsNoEmpty()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Non-empty title",
                TextShapes = [new TestTextShapeDefinition { Paragraphs = ["Some content"] }]
            });

        var result = Service.FindEmptyShapes(path);

        Assert.True(result.Success);
        Assert.Empty(result.EmptyShapes);
    }

    [Fact]
    public void FindEmptyShapes_EmptyTextShape_ReturnsEmpty()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes = [new TestTextShapeDefinition { Name = "Empty Box", Paragraphs = [] }]
            });

        var result = Service.FindEmptyShapes(path);

        Assert.True(result.Success);
        Assert.Single(result.EmptyShapes);
        Assert.Equal("Empty Box", result.EmptyShapes[0].ShapeName);
    }

    [Fact]
    public void FindEmptyShapes_SlideFilter_SearchesOnlySpecifiedSlide()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition
            {
                TextShapes = [new TestTextShapeDefinition { Name = "Empty Shape", Paragraphs = [] }]
            });

        // Search slide 1 — should find no empty shapes
        var result = Service.FindEmptyShapes(path, slideNumber: 1);

        Assert.True(result.Success);
        Assert.Equal(1, result.SlidesSearched);
        Assert.Empty(result.EmptyShapes);
    }

    [Fact]
    public void FindEmptyShapes_OutOfRangeSlideNumber_ReturnsFailure()
    {
        var path = CreateMinimalPptx("Test");
        var result = Service.FindEmptyShapes(path, slideNumber: 99);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FindEmptyShapes_ReturnsCorrectSlideNumber()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition
            {
                TextShapes = [new TestTextShapeDefinition { Name = "Empty Shape" }]
            });

        var result = Service.FindEmptyShapes(path);

        Assert.True(result.Success);
        var emptyOnSlide2 = result.EmptyShapes.Where(s => s.ShapeName == "Empty Shape").ToList();
        Assert.Single(emptyOnSlide2);
        Assert.Equal(2, emptyOnSlide2[0].SlideNumber);
    }
}
