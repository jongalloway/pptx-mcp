using ModelContextProtocol.Protocol;
using PptxTools.Completions;

namespace PptxTools.Tests.Completions;

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

    // --- Action completions ---

    [Fact]
    public void GetCompletions_Action_EmptyPartial_ReturnsAllActions()
    {
        var result = Invoke("action");
        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("Add", result.Completion.Values);
        Assert.Contains("Read", result.Completion.Values);
        Assert.Contains("Update", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_Action_PartialFilter_FiltersCorrectly()
    {
        var result = Invoke("action", "An");
        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("An", v, StringComparison.OrdinalIgnoreCase));
        Assert.Contains("Analyze", result.Completion.Values);
        Assert.Contains("AnalyzeVideo", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_Action_CaseInsensitive()
    {
        var result = Invoke("ACTION", "add");
        Assert.Contains("Add", result.Completion.Values);
    }

    // --- Format completions ---

    [Fact]
    public void GetCompletions_Format_EmptyPartial_ReturnsAllFormats()
    {
        var result = Invoke("format");
        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("png", result.Completion.Values);
        Assert.Contains("markdown", result.Completion.Values);
        Assert.Contains("html", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_Format_PartialFilter_FiltersCorrectly()
    {
        var result = Invoke("format", "jp");
        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("jp", v, StringComparison.OrdinalIgnoreCase));
        Assert.Contains("jpg", result.Completion.Values);
        Assert.Contains("jpeg", result.Completion.Values);
    }

    // --- Style completions ---

    [Fact]
    public void GetCompletions_Style_EmptyPartial_ReturnsAllStyles()
    {
        var result = Invoke("style");
        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("bullet-points", result.Completion.Values);
        Assert.Contains("narrative", result.Completion.Values);
        Assert.Contains("timing-cues", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_Style_PartialFilter_FiltersCorrectly()
    {
        var result = Invoke("style", "bu");
        Assert.Single(result.Completion.Values);
        Assert.Contains("bullet-points", result.Completion.Values);
    }

    // --- ChartAction completions ---

    [Fact]
    public void GetCompletions_ChartAction_EmptyPartial_ReturnsAllChartActions()
    {
        var result = Invoke("chartAction");
        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("Read", result.Completion.Values);
        Assert.Contains("Update", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_ChartAction_PartialFilter_FiltersCorrectly()
    {
        var result = Invoke("chartAction", "Re");
        Assert.Single(result.Completion.Values);
        Assert.Contains("Read", result.Completion.Values);
    }

    // --- SlideNumber completions ---

    [Fact]
    public void GetCompletions_SlideNumber_WithFileContext_ReturnsSlideNumbers()
    {
        var path = CreateTempPptx(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" },
            new TestSlideDefinition { TitleText = "Slide 3" });

        var result = Invoke(
            "slideNumber",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.Equal(3, result.Completion.Values.Count);
        Assert.Contains("1", result.Completion.Values);
        Assert.Contains("2", result.Completion.Values);
        Assert.Contains("3", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_SlideIndex_AliasWorks()
    {
        var path = CreateTempPptx(
            new TestSlideDefinition { TitleText = "Slide 1" });

        var result = Invoke(
            "slideIndex",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.Single(result.Completion.Values);
        Assert.Contains("1", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_SlideNumber_NoFileContext_ReturnsEmpty()
    {
        var result = Invoke("slideNumber", service: Service);
        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_SlideNumber_PartialFilter_FiltersCorrectly()
    {
        var slides = Enumerable.Range(1, 12)
            .Select(i => new TestSlideDefinition { TitleText = $"Slide {i}" })
            .ToArray();
        var path = CreateTempPptx(slides);

        var result = Invoke(
            "slideNumber",
            partialValue: "1",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("1", v, StringComparison.OrdinalIgnoreCase));
        Assert.Contains("1", result.Completion.Values);
        Assert.Contains("10", result.Completion.Values);
        Assert.Contains("11", result.Completion.Values);
        Assert.Contains("12", result.Completion.Values);
    }

    // --- TableName completions ---

    [Fact]
    public void GetCompletions_TableName_WithFileContext_ReturnsTableNames()
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Data Slide",
            Tables =
            [
                new TestTableDefinition { Name = "Revenue Table", Rows = [["Q1", "Q2"], ["100", "200"]] },
                new TestTableDefinition { Name = "Cost Table", Rows = [["Q1", "Q2"], ["50", "75"]] }
            ]
        });

        var result = Invoke(
            "tableName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("Revenue Table", result.Completion.Values);
        Assert.Contains("Cost Table", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_Table_AliasWorks()
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Data Slide",
            Tables = [new TestTableDefinition { Name = "My Table", Rows = [["A", "B"]] }]
        });

        var result = Invoke(
            "table",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.Contains("My Table", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_TableName_NoFileContext_ReturnsEmpty()
    {
        var result = Invoke("tableName", service: Service);
        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_TableName_PartialFilter_FiltersCorrectly()
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Data Slide",
            Tables =
            [
                new TestTableDefinition { Name = "Revenue Table", Rows = [["Q1"], ["100"]] },
                new TestTableDefinition { Name = "Cost Table", Rows = [["Q1"], ["50"]] }
            ]
        });

        var result = Invoke(
            "tableName",
            partialValue: "Rev",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.Single(result.Completion.Values);
        Assert.Contains("Revenue Table", result.Completion.Values);
        Assert.DoesNotContain("Cost Table", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_TableName_ExcludesNonTableShapes()
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Mixed Slide",
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "Text Shape", Paragraphs = ["hello"] }
            ],
            Tables =
            [
                new TestTableDefinition { Name = "Data Table", Rows = [["A"]] }
            ]
        });

        var result = Invoke(
            "tableName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.Contains("Data Table", result.Completion.Values);
        Assert.DoesNotContain("Text Shape", result.Completion.Values);
    }

    [Fact]
    public void GetCompletions_TableName_DeduplicatesAcrossSlides()
    {
        var path = CreateTempPptx(
            new TestSlideDefinition
            {
                TitleText = "Slide 1",
                Tables = [new TestTableDefinition { Name = "Shared Table", Rows = [["A"]] }]
            },
            new TestSlideDefinition
            {
                TitleText = "Slide 2",
                Tables = [new TestTableDefinition { Name = "Shared Table", Rows = [["B"]] }]
            });

        var result = Invoke(
            "tableName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        var count = result.Completion.Values.Count(v => v.Equals("Shared Table", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(1, count);
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
