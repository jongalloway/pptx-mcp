using ModelContextProtocol.Protocol;
using PptxMcp.Completions;

namespace PptxMcp.Tests.Completions;

[Trait("Category", "Integration")]
public class PptxCompletionHandlerTests : PptxTestBase
{
    private string CreateTempPptx(params TestSlideDefinition[] slides)
    {
        if (slides.Length > 0)
            return CreatePptxWithSlides(slides);
        return CreateMinimalPptx();
    }

    private static CompleteResult Invoke(
        string? argumentName,
        string partialValue = "",
        PresentationService? service = null,
        Dictionary<string, string>? contextArgs = null)
        => PptxCompletionHandler.GetCompletions(argumentName, partialValue, contextArgs, service);

    // --- Null / empty argument name tests ---

    [Fact]
    public void GetCompletions_NullArgumentName_ReturnsEmpty()
    {
        var result = Invoke(null);
        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_EmptyArgumentName_ReturnsEmpty()
    {
        var result = Invoke(string.Empty);
        Assert.Empty(result.Completion.Values);
    }

    // --- PlaceholderType completions ---

    [Fact]
    public void GetCompletions_PlaceholderType_EmptyPartial_ReturnsAllTypes()
    {
        var result = Invoke("placeholderType");
        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("title", result.Completion.Values);
        Assert.Contains("body", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_PlaceholderType_PartialTitle_FiltersCorrectly()
    {
        var result = Invoke("placeholderType", "ti");
        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("ti", v, StringComparison.OrdinalIgnoreCase));
        Assert.Contains("title", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_PlaceholderType_NonMatchingPartial_ReturnsEmpty()
    {
        var result = Invoke("placeholderType", "zzz");
        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_PlaceholderType_CaseInsensitive()
    {
        var result = Invoke("placeholderType", "TITLE");
        Assert.Contains("title", result.Completion.Values);
    }

    // --- LayoutName completions ---

    [Fact]
    public void GetCompletions_LayoutName_WithFileContext_ReturnsLayouts()
    {
        var path = CreateTempPptx();
        var result = Invoke(
            "layoutName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("Title Slide", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_LayoutName_WithPartial_FiltersCorrectly()
    {
        var path = CreateTempPptx();
        var result = Invoke(
            "layoutName",
            partialValue: "Title",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("Title", v, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GetCompletions_LayoutName_NoFileContext_ReturnsEmpty()
    {
        var result = Invoke("layoutName", service: Service);
        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_LayoutName_FileNotFound_ReturnsEmpty()
    {
        var result = Invoke(
            "layoutName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = "/nonexistent/deck.pptx" });

        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_LayoutName_UrlEncodedFilePath_Decodes()
    {
        var path = CreateTempPptx();
        var encoded = Uri.EscapeDataString(path);
        var result = Invoke(
            "layoutName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = encoded });

        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_LayoutName_NullService_ReturnsEmpty()
    {
        var path = CreateTempPptx();
        var result = Invoke(
            "layoutName",
            service: null,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.Empty(result.Completion.Values);
    }

    // --- ShapeName completions ---

    [Fact]
    public void GetCompletions_ShapeName_WithFileContext_ReturnsShapeNames()
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Deck Title",
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "Revenue Value", Paragraphs = ["$1M"] },
                new TestTextShapeDefinition { Name = "Growth Rate", Paragraphs = ["12%"] }
            ]
        });

        var result = Invoke(
            "shapeName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("Revenue Value", result.Completion.Values);
        Assert.Contains("Growth Rate", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_ShapeName_WithPartial_FiltersCorrectly()
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Deck Title",
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "Revenue Value", Paragraphs = ["$1M"] },
                new TestTextShapeDefinition { Name = "Growth Rate", Paragraphs = ["12%"] }
            ]
        });

        var result = Invoke(
            "shapeName",
            partialValue: "Rev",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("Rev", v, StringComparison.OrdinalIgnoreCase));
        Assert.Contains("Revenue Value", result.Completion.Values);
        Assert.DoesNotContain("Growth Rate", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_ShapeName_AcceptsFilePathKey()
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Title",
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "My Shape", Paragraphs = ["text"] }
            ]
        });

        var result = Invoke(
            "shapeName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["filePath"] = path });

        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("My Shape", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_ShapeName_NoFileContext_ReturnsEmpty()
    {
        var result = Invoke("shapeName", service: Service);
        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_ShapeName_DeduplicatesAcrossSlides()
    {
        var sharedName = "Shared Shape";
        var path = CreateTempPptx(
            new TestSlideDefinition
            {
                TitleText = "Slide 1",
                TextShapes =
                [
                    new TestTextShapeDefinition { Name = sharedName, Paragraphs = ["value1"] }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Slide 2",
                TextShapes =
                [
                    new TestTextShapeDefinition { Name = sharedName, Paragraphs = ["value2"] }
                ]
            });

        var result = Invoke(
            "shapeName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        var count = result.Completion.Values.Count(v => v.Equals(sharedName, StringComparison.OrdinalIgnoreCase));
        Assert.Equal(1, count);
    }

    // --- Unknown argument name ---

    [Fact]
    public void GetCompletions_UnknownArgumentName_ReturnsEmpty()
    {
        var result = Invoke("unknownArgument");
        Assert.Empty(result.Completion.Values);
    }

    // --- Result structure ---

    [Fact]
    public void GetCompletions_Result_HasCorrectTotalCount()
    {
        var result = Invoke("placeholderType");
        Assert.Equal(result.Completion.Values.Count, result.Completion.Total);
    }

    [Fact]
    public void GetCompletions_Result_HasMoreIsFalse()
    {
        var result = Invoke("placeholderType");
        Assert.False(result.Completion.HasMore);
    }
}
