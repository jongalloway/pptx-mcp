using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests.Tools;

[Trait("Category", "Integration")]
public class TextFormattingToolsTests : PptxTestBase
{
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public TextFormattingToolsTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    // ── Fixture helpers ──────────────────────────────────────────────────────────

    private string CreateFormattedPptx(
        string shapeName = "Styled",
        string text = "Sample",
        string? fontFamily = null,
        int? fontSizeHundredths = null,
        bool? bold = null,
        string? colorHex = null,
        A.TextAlignmentTypeValues? alignment = null)
    {
        var path = CreateMinimalPptx("Tool Test Slide");

        using var doc = PresentationDocument.Open(path, true);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var tree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        var runProps = new A.RunProperties();
        if (bold.HasValue) runProps.Bold = bold.Value;
        if (fontSizeHundredths.HasValue) runProps.FontSize = fontSizeHundredths.Value;
        if (fontFamily is not null) runProps.Append(new A.LatinFont { Typeface = fontFamily });
        if (colorHex is not null)
            runProps.InsertAt(new A.SolidFill(new A.RgbColorModelHex { Val = colorHex }), 0);

        var paragraphProps = alignment.HasValue
            ? new A.ParagraphProperties { Alignment = alignment.Value }
            : null;

        var paragraph = new A.Paragraph();
        if (paragraphProps is not null) paragraph.Append(paragraphProps);
        paragraph.Append(new A.Run(runProps, new A.Text(text)));
        paragraph.Append(new A.EndParagraphRunProperties());

        var textBody = new TextBody(new A.BodyProperties(), new A.ListStyle(), paragraph);

        var nextId = (uint)(tree.Elements<Shape>().Count() + 10);
        tree.Append(new Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = nextId, Name = shapeName },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 100000, Y = 100000 },
                    new A.Extents { Cx = 5000000, Cy = 1000000 })),
            textBody));

        slidePart.Slide.Save();
        return path;
    }

    // ── Get action: structured JSON ──────────────────────────────────────────────

    [Fact]
    public async Task Get_ReturnsStructuredJsonWithFormattingInfo()
    {
        var path = CreateFormattedPptx(
            shapeName: "StyledBox",
            text: "Hello",
            fontFamily: "Arial",
            fontSizeHundredths: 2000,
            bold: true,
            colorHex: "FF0000",
            alignment: A.TextAlignmentTypeValues.Center);

        var result = await _tools.pptx_manage_text_formatting(path, TextFormattingAction.Get);

        var parsed = JsonSerializer.Deserialize<TextFormattingResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("Get", parsed.Action);
        Assert.True(parsed.FormattingCount > 0);

        var run = Assert.Single(parsed.Formattings, f => f.ShapeName == "StyledBox");
        Assert.Equal("Hello", run.Text);
        Assert.Equal("Arial", run.FontFamily);
        Assert.Equal(20.0, run.FontSize);
        Assert.True(run.Bold);
        Assert.Equal("#FF0000", run.Color);
        Assert.Equal("Center", run.Alignment);
    }

    [Fact]
    public async Task Get_WithSlideNumberFilter_ReturnsFilteredResults()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes = [new TestTextShapeDefinition { Name = "S1Shape", Paragraphs = ["Slide1 text"] }]
            },
            new TestSlideDefinition
            {
                TextShapes = [new TestTextShapeDefinition { Name = "S2Shape", Paragraphs = ["Slide2 text"] }]
            });

        var result = await _tools.pptx_manage_text_formatting(path, TextFormattingAction.Get, slideNumber: 2);

        var parsed = JsonSerializer.Deserialize<TextFormattingResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.All(parsed.Formattings, f => Assert.Equal(2, f.SlideNumber));
    }

    [Fact]
    public async Task Get_WithShapeNameFilter_ReturnsOnlyMatchingShape()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "Alpha", Paragraphs = ["First"] },
                new TestTextShapeDefinition { Name = "Beta", Paragraphs = ["Second"] }
            ]
        });

        var result = await _tools.pptx_manage_text_formatting(path, TextFormattingAction.Get, shapeName: "Alpha");

        var parsed = JsonSerializer.Deserialize<TextFormattingResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.All(parsed.Formattings, f => Assert.Equal("Alpha", f.ShapeName));
    }

    // ── Get action: file not found ───────────────────────────────────────────────

    [Fact]
    public async Task Get_FileNotFound_ReturnsStructuredError()
    {
        var fakePath = Path.Join(Path.GetTempPath(), "nonexistent-tf-tool.pptx");

        var result = await _tools.pptx_manage_text_formatting(fakePath, TextFormattingAction.Get);

        var parsed = JsonSerializer.Deserialize<TextFormattingResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("File not found", parsed.Message);
    }

    // ── Apply action: success ────────────────────────────────────────────────────

    [Fact]
    public async Task Apply_ReturnsSuccessResult()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        var result = await _tools.pptx_manage_text_formatting(path, TextFormattingAction.Apply,
            slideNumber: 1, shapeName: "Target", bold: true, fontSize: 24.0);

        var parsed = JsonSerializer.Deserialize<TextFormattingResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("Apply", parsed.Action);
        Assert.True(parsed.FormattingCount > 0);
        Assert.Contains("Applied formatting", parsed.Message);
    }

    [Fact]
    public async Task Apply_ChangesArePersistedAndReadable()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        await _tools.pptx_manage_text_formatting(path, TextFormattingAction.Apply,
            slideNumber: 1, shapeName: "Target",
            fontFamily: "Consolas", fontSize: 12.0, bold: true, italic: true,
            color: "#AABBCC", alignment: "Right");

        var getResult = await _tools.pptx_manage_text_formatting(path, TextFormattingAction.Get,
            shapeName: "Target");
        var parsed = JsonSerializer.Deserialize<TextFormattingResult>(getResult);
        Assert.NotNull(parsed);
        var run = Assert.Single(parsed.Formattings, f => f.ShapeName == "Target");
        Assert.Equal("Consolas", run.FontFamily);
        Assert.Equal(12.0, run.FontSize);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.Equal("#AABBCC", run.Color);
        Assert.Equal("Right", run.Alignment);
    }

    // ── Apply action: required params ────────────────────────────────────────────

    [Fact]
    public async Task Apply_MissingSlideNumber_ReturnsStructuredError()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        var result = await _tools.pptx_manage_text_formatting(path, TextFormattingAction.Apply,
            slideNumber: null, shapeName: "Target", bold: true);

        var parsed = JsonSerializer.Deserialize<TextFormattingResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("slideNumber is required", parsed.Message);
    }

    [Fact]
    public async Task Apply_MissingShapeName_ReturnsStructuredError()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        var result = await _tools.pptx_manage_text_formatting(path, TextFormattingAction.Apply,
            slideNumber: 1, shapeName: null, bold: true);

        var parsed = JsonSerializer.Deserialize<TextFormattingResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("shapeName is required", parsed.Message);
    }

    // ── Apply action: file not found ─────────────────────────────────────────────

    [Fact]
    public async Task Apply_FileNotFound_ReturnsStructuredError()
    {
        var fakePath = Path.Join(Path.GetTempPath(), "nonexistent-apply-tool.pptx");

        var result = await _tools.pptx_manage_text_formatting(fakePath, TextFormattingAction.Apply,
            slideNumber: 1, shapeName: "Target", bold: true);

        var parsed = JsonSerializer.Deserialize<TextFormattingResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("File not found", parsed.Message);
    }

    // ── Apply action: shape/slide not found ──────────────────────────────────────

    [Fact]
    public async Task Apply_ShapeNotFound_ReturnsStructuredError()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_manage_text_formatting(path, TextFormattingAction.Apply,
            slideNumber: 1, shapeName: "GhostShape", bold: true);

        var parsed = JsonSerializer.Deserialize<TextFormattingResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("No shape named", parsed.Message);
    }

    [Fact]
    public async Task Apply_SlideOutOfRange_ReturnsStructuredError()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_manage_text_formatting(path, TextFormattingAction.Apply,
            slideNumber: 99, shapeName: "Target", bold: true);

        var parsed = JsonSerializer.Deserialize<TextFormattingResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("out of range", parsed.Message);
    }

    // ── Unknown action ───────────────────────────────────────────────────────────

    [Fact]
    public async Task UnknownAction_ReturnsErrorJson()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_manage_text_formatting(path, (TextFormattingAction)99);

        Assert.Contains("Unknown action", result);
    }
}
