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

    // ================================================================
    // Null / empty argument name
    // ================================================================

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

    // ================================================================
    // PlaceholderType completions (existing)
    // ================================================================

    [Fact]
    public void PlaceholderType_EmptyPartial_ReturnsAllTypes()
    {
        var result = Invoke("placeholderType");
        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("title", result.Completion.Values);
        Assert.Contains("body", result.Completion.Values);
    }

    [Fact]
    public void PlaceholderType_WithPartialValue_FiltersResults()
    {
        var result = Invoke("placeholderType", "ti");
        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("ti", v, StringComparison.OrdinalIgnoreCase));
        Assert.Contains("title", result.Completion.Values);
    }

    [Fact]
    public void PlaceholderType_NonMatchingPartial_ReturnsEmpty()
    {
        var result = Invoke("placeholderType", "zzz");
        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void PlaceholderType_CaseInsensitiveArgumentName()
    {
        var result = Invoke("PLACEHOLDERTYPE");
        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("title", result.Completion.Values);
    }

    [Fact]
    public void PlaceholderType_CaseInsensitiveFiltering()
    {
        var result = Invoke("placeholderType", "TITLE");
        Assert.Contains("title", result.Completion.Values);
    }

    [Fact]
    public void PlaceholderType_ContainsExpectedValues()
    {
        var result = Invoke("placeholderType");
        var values = result.Completion.Values;
        Assert.Contains("ctrTitle", values);
        Assert.Contains("subTitle", values);
        Assert.Contains("dt", values);
        Assert.Contains("ftr", values);
        Assert.Contains("sldNum", values);
        Assert.Contains("obj", values);
        Assert.Contains("tbl", values);
        Assert.Contains("chart", values);
    }

    // ================================================================
    // LayoutName completions (existing)
    // ================================================================

    [Fact]
    public void LayoutName_WithFileContext_ReturnsLayouts()
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
    public void LayoutName_WithPartialValue_FiltersResults()
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
    public void LayoutName_NoFileContext_ReturnsEmpty()
    {
        var result = Invoke("layoutName", service: Service);
        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void LayoutName_FileNotFound_ReturnsEmpty()
    {
        var result = Invoke(
            "layoutName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = "/nonexistent/deck.pptx" });

        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void LayoutName_UrlEncodedFilePath_Decodes()
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
    public void LayoutName_NullService_ReturnsEmpty()
    {
        var path = CreateTempPptx();
        var result = Invoke(
            "layoutName",
            service: null,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void Layout_AliasCaseInsensitive()
    {
        var path = CreateTempPptx();
        var result = Invoke(
            "layout",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
    }

    // ================================================================
    // ShapeName completions (existing)
    // ================================================================

    [Fact]
    public void ShapeName_WithFileContext_ReturnsShapeNames()
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
    public void ShapeName_WithPartialValue_FiltersResults()
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
    public void ShapeName_AcceptsFilePathKey()
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
    public void ShapeName_NoFileContext_ReturnsEmpty()
    {
        var result = Invoke("shapeName", service: Service);
        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void ShapeName_DeduplicatesAcrossSlides()
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

    [Fact]
    public void Shape_AliasCaseInsensitive()
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Title",
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "TestShape", Paragraphs = ["x"] }
            ]
        });

        var result = Invoke(
            "shape",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
    }

    // ================================================================
    // Action completions (new - static)
    // ================================================================

    [Fact]
    public void Action_ReturnsNonEmptyCompletions()
    {
        var result = Invoke("action");
        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void Action_CaseInsensitiveArgumentName()
    {
        var result = Invoke("ACTION");
        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void Action_EmptyPartial_ReturnsAllValues()
    {
        var result = Invoke("action", "");
        Assert.True(result.Completion.Values.Count >= 10,
            "Expected at least 10 action completions");
    }

    [Fact]
    public void Action_WithPartialValue_FiltersResults()
    {
        var result = Invoke("action", "Add");
        Assert.NotEmpty(result.Completion.Values);
        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("Add", v, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Action_ContainsExpectedValues()
    {
        var result = Invoke("action");
        var values = result.Completion.Values;
        Assert.Contains("Add", values);
        Assert.Contains("AddFromLayout", values);
        Assert.Contains("Duplicate", values);
        Assert.Contains("Move", values);
        Assert.Contains("Reorder", values);
        Assert.Contains("Find", values);
        Assert.Contains("Remove", values);
        Assert.Contains("Analyze", values);
        Assert.Contains("Deduplicate", values);
        Assert.Contains("AnalyzeVideo", values);
        Assert.Contains("Read", values);
        Assert.Contains("Update", values);
    }

    [Fact]
    public void Action_PartialRe_FiltersToRelevantValues()
    {
        var result = Invoke("action", "Re");
        Assert.NotEmpty(result.Completion.Values);
        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("Re", v, StringComparison.OrdinalIgnoreCase));
        Assert.Contains("Reorder", result.Completion.Values);
        Assert.Contains("Remove", result.Completion.Values);
        Assert.Contains("Read", result.Completion.Values);
    }

    [Fact]
    public void Action_NonMatchingPartial_ReturnsEmpty()
    {
        var result = Invoke("action", "zzz");
        Assert.Empty(result.Completion.Values);
    }

    // ================================================================
    // Format completions (new - static)
    // ================================================================

    [Fact]
    public void Format_ReturnsNonEmptyCompletions()
    {
        var result = Invoke("format");
        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void Format_CaseInsensitiveArgumentName()
    {
        var result = Invoke("FORMAT");
        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void Format_ContainsExpectedValues()
    {
        var result = Invoke("format");
        var values = result.Completion.Values;
        Assert.Contains("png", values);
        Assert.Contains("jpg", values);
        Assert.Contains("jpeg", values);
        Assert.Contains("gif", values);
        Assert.Contains("bmp", values);
        Assert.Contains("tiff", values);
        Assert.Contains("svg", values);
        Assert.Contains("markdown", values);
        Assert.Contains("html", values);
    }

    [Fact]
    public void Format_WithPartialValue_FiltersResults()
    {
        var result = Invoke("format", "j");
        Assert.NotEmpty(result.Completion.Values);
        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("j", v, StringComparison.OrdinalIgnoreCase));
        Assert.Contains("jpg", result.Completion.Values);
        Assert.Contains("jpeg", result.Completion.Values);
    }

    [Fact]
    public void Format_EmptyPartial_ReturnsAllValues()
    {
        var result = Invoke("format", "");
        Assert.Equal(9, result.Completion.Values.Count);
    }

    [Fact]
    public void Format_NonMatchingPartial_ReturnsEmpty()
    {
        var result = Invoke("format", "zzz");
        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void Format_PartialMark_FiltersToMarkdown()
    {
        var result = Invoke("format", "mark");
        Assert.Single(result.Completion.Values);
        Assert.Contains("markdown", result.Completion.Values);
    }

    // ================================================================
    // Style completions (new - static)
    // ================================================================

    [Fact]
    public void Style_ReturnsNonEmptyCompletions()
    {
        var result = Invoke("style");
        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void Style_CaseInsensitiveArgumentName()
    {
        var result = Invoke("STYLE");
        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void Style_ContainsExpectedValues()
    {
        var result = Invoke("style");
        var values = result.Completion.Values;
        Assert.Contains("bullet-points", values);
        Assert.Contains("narrative", values);
        Assert.Contains("timing-cues", values);
    }

    [Fact]
    public void Style_EmptyPartial_ReturnsAllValues()
    {
        var result = Invoke("style", "");
        Assert.Equal(3, result.Completion.Values.Count);
    }

    [Fact]
    public void Style_WithPartialValue_FiltersResults()
    {
        var result = Invoke("style", "bullet");
        Assert.Single(result.Completion.Values);
        Assert.Contains("bullet-points", result.Completion.Values);
    }

    [Fact]
    public void Style_NonMatchingPartial_ReturnsEmpty()
    {
        var result = Invoke("style", "zzz");
        Assert.Empty(result.Completion.Values);
    }

    [Fact]
    public void Style_PartialT_FiltersToTimingCues()
    {
        var result = Invoke("style", "t");
        Assert.Single(result.Completion.Values);
        Assert.Contains("timing-cues", result.Completion.Values);
    }

    // ================================================================
    // ChartAction completions (new - static)
    // ================================================================

    [Fact]
    public void ChartAction_ReturnsNonEmptyCompletions()
    {
        var result = Invoke("chartAction");
        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void ChartAction_CaseInsensitiveArgumentName()
    {
        var result = Invoke("CHARTACTION");
        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void ChartAction_ContainsExpectedValues()
    {
        var result = Invoke("chartAction");
        var values = result.Completion.Values;
        Assert.Contains("Read", values);
        Assert.Contains("Update", values);
    }

    [Fact]
    public void ChartAction_EmptyPartial_ReturnsAllValues()
    {
        var result = Invoke("chartAction", "");
        Assert.Equal(2, result.Completion.Values.Count);
    }

    [Fact]
    public void ChartAction_WithPartialValue_FiltersResults()
    {
        var result = Invoke("chartAction", "R");
        Assert.Single(result.Completion.Values);
        Assert.Contains("Read", result.Completion.Values);
    }

    [Fact]
    public void ChartAction_NonMatchingPartial_ReturnsEmpty()
    {
        var result = Invoke("chartAction", "zzz");
        Assert.Empty(result.Completion.Values);
    }

    // ================================================================
    // SlideNumber / SlideIndex completions (new - dynamic)
    // ================================================================

    [Theory]
    [InlineData("slideNumber")]
    [InlineData("slideIndex")]
    public void SlideNumber_WithoutService_ReturnsEmpty(string argName)
    {
        var result = Invoke(argName, service: null);
        Assert.Empty(result.Completion.Values);
    }

    [Theory]
    [InlineData("slideNumber")]
    [InlineData("slideIndex")]
    public void SlideNumber_WithoutFilePath_ReturnsEmpty(string argName)
    {
        var result = Invoke(argName, service: Service);
        Assert.Empty(result.Completion.Values);
    }

    [Theory]
    [InlineData("slideNumber")]
    [InlineData("slideIndex")]
    public void SlideNumber_FileNotFound_ReturnsEmpty(string argName)
    {
        var result = Invoke(
            argName,
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = "/nonexistent/deck.pptx" });

        Assert.Empty(result.Completion.Values);
    }

    [Theory]
    [InlineData("slideNumber")]
    [InlineData("slideIndex")]
    public void SlideNumber_WithValidFile_ReturnsSlideNumbers(string argName)
    {
        var path = CreateTempPptx(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" },
            new TestSlideDefinition { TitleText = "Slide 3" });

        var result = Invoke(
            argName,
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
        Assert.True(result.Completion.Values.Count >= 3,
            "Expected at least 3 slide number completions for a 3-slide deck");
    }

    [Theory]
    [InlineData("slideNumber")]
    [InlineData("slideIndex")]
    public void SlideNumber_SingleSlide_ReturnsSingleValue(string argName)
    {
        var path = CreateTempPptx();
        var result = Invoke(
            argName,
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
    }

    [Theory]
    [InlineData("slideNumber")]
    [InlineData("slideIndex")]
    public void SlideNumber_CaseInsensitiveArgumentName(string argName)
    {
        var path = CreateTempPptx();
        var result = Invoke(
            argName.ToUpperInvariant(),
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void SlideNumber_WithPartialValue_FiltersResults()
    {
        var slides = Enumerable.Range(1, 12)
            .Select(i => new TestSlideDefinition { TitleText = $"Slide {i}" })
            .ToArray();
        var path = CreatePptxWithSlides(slides);

        var result = Invoke(
            "slideNumber",
            partialValue: "1",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("1", v, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void SlideNumber_AcceptsFilePathKey()
    {
        var path = CreateTempPptx(
            new TestSlideDefinition { TitleText = "S1" },
            new TestSlideDefinition { TitleText = "S2" });

        var result = Invoke(
            "slideNumber",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["filePath"] = path });

        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void SlideNumber_UrlEncodedFilePath_Decodes()
    {
        var path = CreateTempPptx();
        var encoded = Uri.EscapeDataString(path);
        var result = Invoke(
            "slideNumber",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = encoded });

        Assert.NotEmpty(result.Completion.Values);
    }

    // ================================================================
    // TableName / Table completions (new - dynamic)
    // ================================================================

    [Theory]
    [InlineData("tableName")]
    [InlineData("table")]
    public void TableName_WithoutService_ReturnsEmpty(string argName)
    {
        var result = Invoke(argName, service: null);
        Assert.Empty(result.Completion.Values);
    }

    [Theory]
    [InlineData("tableName")]
    [InlineData("table")]
    public void TableName_WithoutFilePath_ReturnsEmpty(string argName)
    {
        var result = Invoke(argName, service: Service);
        Assert.Empty(result.Completion.Values);
    }

    [Theory]
    [InlineData("tableName")]
    [InlineData("table")]
    public void TableName_FileNotFound_ReturnsEmpty(string argName)
    {
        var result = Invoke(
            argName,
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = "/nonexistent/deck.pptx" });

        Assert.Empty(result.Completion.Values);
    }

    [Theory]
    [InlineData("tableName")]
    [InlineData("table")]
    public void TableName_WithValidFile_ReturnsTableNames(string argName)
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Data Slide",
            Tables =
            [
                new TestTableDefinition
                {
                    Name = "Revenue Table",
                    Rows = [["Q1", "100"], ["Q2", "200"]]
                },
                new TestTableDefinition
                {
                    Name = "Cost Table",
                    Rows = [["Q1", "50"], ["Q2", "75"]]
                }
            ]
        });

        var result = Invoke(
            argName,
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("Revenue Table", result.Completion.Values);
        Assert.Contains("Cost Table", result.Completion.Values);
    }

    [Fact]
    public void TableName_WithPartialValue_FiltersResults()
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Data Slide",
            Tables =
            [
                new TestTableDefinition
                {
                    Name = "Revenue Table",
                    Rows = [["Q1", "100"]]
                },
                new TestTableDefinition
                {
                    Name = "Cost Table",
                    Rows = [["Q1", "50"]]
                }
            ]
        });

        var result = Invoke(
            "tableName",
            partialValue: "Rev",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
        Assert.All(result.Completion.Values,
            v => Assert.StartsWith("Rev", v, StringComparison.OrdinalIgnoreCase));
        Assert.Contains("Revenue Table", result.Completion.Values);
        Assert.DoesNotContain("Cost Table", result.Completion.Values);
    }

    [Fact]
    public void TableName_AcrossMultipleSlides_ReturnsAll()
    {
        var path = CreateTempPptx(
            new TestSlideDefinition
            {
                TitleText = "Slide 1",
                Tables =
                [
                    new TestTableDefinition
                    {
                        Name = "Table Alpha",
                        Rows = [["A", "B"]]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Slide 2",
                Tables =
                [
                    new TestTableDefinition
                    {
                        Name = "Table Beta",
                        Rows = [["C", "D"]]
                    }
                ]
            });

        var result = Invoke(
            "tableName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("Table Alpha", result.Completion.Values);
        Assert.Contains("Table Beta", result.Completion.Values);
    }

    [Fact]
    public void TableName_DeduplicatesAcrossSlides()
    {
        var path = CreateTempPptx(
            new TestSlideDefinition
            {
                TitleText = "Slide 1",
                Tables =
                [
                    new TestTableDefinition
                    {
                        Name = "Shared Table",
                        Rows = [["A", "B"]]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Slide 2",
                Tables =
                [
                    new TestTableDefinition
                    {
                        Name = "Shared Table",
                        Rows = [["C", "D"]]
                    }
                ]
            });

        var result = Invoke(
            "tableName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        var count = result.Completion.Values.Count(
            v => v.Equals("Shared Table", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(1, count);
    }

    [Fact]
    public void TableName_NoTables_ReturnsEmpty()
    {
        var path = CreateTempPptx(new TestSlideDefinition { TitleText = "No tables here" });
        var result = Invoke(
            "tableName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.Empty(result.Completion.Values);
    }

    [Theory]
    [InlineData("tableName")]
    [InlineData("table")]
    public void TableName_CaseInsensitiveArgumentName(string argName)
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Data",
            Tables =
            [
                new TestTableDefinition
                {
                    Name = "TestTable",
                    Rows = [["X", "Y"]]
                }
            ]
        });

        var result = Invoke(
            argName.ToUpperInvariant(),
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.NotEmpty(result.Completion.Values);
    }

    [Fact]
    public void TableName_AcceptsFilePathKey()
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Data",
            Tables =
            [
                new TestTableDefinition
                {
                    Name = "MyTable",
                    Rows = [["A", "B"]]
                }
            ]
        });

        var result = Invoke(
            "tableName",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["filePath"] = path });

        Assert.NotEmpty(result.Completion.Values);
        Assert.Contains("MyTable", result.Completion.Values);
    }

    [Fact]
    public void TableName_NonMatchingPartial_ReturnsEmpty()
    {
        var path = CreateTempPptx(new TestSlideDefinition
        {
            TitleText = "Data",
            Tables =
            [
                new TestTableDefinition
                {
                    Name = "Revenue",
                    Rows = [["Q1", "100"]]
                }
            ]
        });

        var result = Invoke(
            "tableName",
            partialValue: "zzz",
            service: Service,
            contextArgs: new Dictionary<string, string> { ["file"] = path });

        Assert.Empty(result.Completion.Values);
    }

    // ================================================================
    // Unknown argument name
    // ================================================================

    [Fact]
    public void GetCompletions_UnknownArgumentName_ReturnsEmpty()
    {
        var result = Invoke("unknownArgument");
        Assert.Empty(result.Completion.Values);
    }

    // ================================================================
    // Result structure invariants
    // ================================================================

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

    [Fact]
    public void Action_Result_TotalMatchesValuesCount()
    {
        var result = Invoke("action");
        Assert.Equal(result.Completion.Values.Count, result.Completion.Total);
    }

    [Fact]
    public void Format_Result_TotalMatchesValuesCount()
    {
        var result = Invoke("format");
        Assert.Equal(result.Completion.Values.Count, result.Completion.Total);
    }

    [Fact]
    public void Style_Result_HasMoreIsFalse()
    {
        var result = Invoke("style");
        Assert.False(result.Completion.HasMore);
    }
}