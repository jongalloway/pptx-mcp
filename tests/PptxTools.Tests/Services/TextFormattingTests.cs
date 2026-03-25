using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests.Services;

[Trait("Category", "Unit")]
public class TextFormattingTests : PptxTestBase
{
    // ── Fixture helpers ──────────────────────────────────────────────────────────

    /// <summary>Creates a PPTX with a single shape whose runs have explicit formatting.</summary>
    private string CreateFormattedPptx(
        string shapeName = "Formatted",
        string text = "Hello",
        string? fontFamily = null,
        int? fontSizeHundredths = null,
        bool? bold = null,
        bool? italic = null,
        A.TextUnderlineValues? underline = null,
        string? colorHex = null,
        A.TextAlignmentTypeValues? alignment = null)
    {
        var path = CreateMinimalPptx("Slide 1");

        using var doc = PresentationDocument.Open(path, true);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var tree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        var runProps = new A.RunProperties();
        if (bold.HasValue) runProps.Bold = bold.Value;
        if (italic.HasValue) runProps.Italic = italic.Value;
        if (underline.HasValue) runProps.Underline = underline.Value;
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

    /// <summary>Creates a PPTX with multiple shapes across multiple slides.</summary>
    private string CreateMultiSlideFormattedPptx()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Slide 1",
                TextShapes =
                [
                    new TestTextShapeDefinition { Name = "Header", Paragraphs = ["Welcome"] },
                    new TestTextShapeDefinition { Name = "Body", Paragraphs = ["Content here"] }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Slide 2",
                TextShapes =
                [
                    new TestTextShapeDefinition { Name = "Footer", Paragraphs = ["Page 2"] }
                ]
            });
        return path;
    }

    /// <summary>Creates a PPTX with a shape that has a text body with no runs.</summary>
    private string CreateShapeWithNoRuns(string shapeName = "EmptyShape")
    {
        var path = CreateMinimalPptx("Slide 1");

        using var doc = PresentationDocument.Open(path, true);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var tree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        var paragraph = new A.Paragraph(new A.EndParagraphRunProperties());
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

    // ── GetTextFormatting: basic retrieval ────────────────────────────────────────

    [Fact]
    public void GetTextFormatting_ReturnsFormattingFromExistingPresentation()
    {
        var path = CreateFormattedPptx(text: "Hello World");

        var result = Service.GetTextFormatting(path);

        Assert.True(result.Success);
        Assert.Equal("Get", result.Action);
        Assert.True(result.FormattingCount > 0);
        var formatted = Assert.Single(result.Formattings, f => f.ShapeName == "Formatted");
        Assert.Equal("Hello World", formatted.Text);
    }

    [Fact]
    public void GetTextFormatting_ReturnsFontFamily()
    {
        var path = CreateFormattedPptx(fontFamily: "Arial");

        var result = Service.GetTextFormatting(path);

        var run = Assert.Single(result.Formattings, f => f.ShapeName == "Formatted");
        Assert.Equal("Arial", run.FontFamily);
    }

    [Fact]
    public void GetTextFormatting_ReturnsFontSizeInPoints()
    {
        // OpenXML stores font size in hundredths of a point; 2400 → 24.0 pt
        var path = CreateFormattedPptx(fontSizeHundredths: 2400);

        var result = Service.GetTextFormatting(path);

        var run = Assert.Single(result.Formattings, f => f.ShapeName == "Formatted");
        Assert.Equal(24.0, run.FontSize);
    }

    [Fact]
    public void GetTextFormatting_ReturnsBoldItalicUnderline()
    {
        var path = CreateFormattedPptx(bold: true, italic: true, underline: A.TextUnderlineValues.Single);

        var result = Service.GetTextFormatting(path);

        var run = Assert.Single(result.Formattings, f => f.ShapeName == "Formatted");
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.True(run.Underline);
    }

    [Fact]
    public void GetTextFormatting_ReturnsColor()
    {
        var path = CreateFormattedPptx(colorHex: "FF0000");

        var result = Service.GetTextFormatting(path);

        var run = Assert.Single(result.Formattings, f => f.ShapeName == "Formatted");
        Assert.Equal("#FF0000", run.Color);
    }

    [Fact]
    public void GetTextFormatting_ReturnsAlignment()
    {
        var path = CreateFormattedPptx(alignment: A.TextAlignmentTypeValues.Center);

        var result = Service.GetTextFormatting(path);

        var run = Assert.Single(result.Formattings, f => f.ShapeName == "Formatted");
        Assert.Equal("Center", run.Alignment);
    }

    [Fact]
    public void GetTextFormatting_NullsWhenNoFormattingSet()
    {
        var path = CreateFormattedPptx(text: "Plain");

        var result = Service.GetTextFormatting(path);

        var run = Assert.Single(result.Formattings, f => f.ShapeName == "Formatted");
        Assert.Null(run.FontFamily);
        Assert.Null(run.FontSize);
        Assert.Null(run.Bold);
        Assert.Null(run.Italic);
        Assert.Null(run.Underline);
        Assert.Null(run.Color);
        Assert.Null(run.Alignment);
    }

    // ── GetTextFormatting: filters ───────────────────────────────────────────────

    [Fact]
    public void GetTextFormatting_WithSlideNumberFilter_ReturnsOnlyThatSlide()
    {
        var path = CreateMultiSlideFormattedPptx();

        var result = Service.GetTextFormatting(path, slideNumber: 2);

        Assert.True(result.Success);
        Assert.All(result.Formattings, f => Assert.Equal(2, f.SlideNumber));
    }

    [Fact]
    public void GetTextFormatting_WithShapeNameFilter_ReturnsOnlyMatchingShape()
    {
        var path = CreateMultiSlideFormattedPptx();

        var result = Service.GetTextFormatting(path, shapeName: "Header");

        Assert.True(result.Success);
        Assert.All(result.Formattings, f => Assert.Equal("Header", f.ShapeName));
    }

    [Fact]
    public void GetTextFormatting_ShapeNameFilter_IsCaseInsensitive()
    {
        var path = CreateMultiSlideFormattedPptx();

        var result = Service.GetTextFormatting(path, shapeName: "HEADER");

        Assert.True(result.Success);
        Assert.True(result.FormattingCount > 0);
        Assert.All(result.Formattings, f => Assert.Equal("Header", f.ShapeName));
    }

    [Fact]
    public void GetTextFormatting_WithBothFilters_ReturnsIntersection()
    {
        var path = CreateMultiSlideFormattedPptx();

        var result = Service.GetTextFormatting(path, slideNumber: 1, shapeName: "Body");

        Assert.True(result.Success);
        Assert.All(result.Formattings, f =>
        {
            Assert.Equal(1, f.SlideNumber);
            Assert.Equal("Body", f.ShapeName);
        });
    }

    [Fact]
    public void GetTextFormatting_AllSlidesReturned_WhenNoSlideFilter()
    {
        var path = CreateMultiSlideFormattedPptx();

        var result = Service.GetTextFormatting(path);

        Assert.True(result.Success);
        var slideNumbers = result.Formattings.Select(f => f.SlideNumber).Distinct().OrderBy(n => n).ToList();
        Assert.Contains(1, slideNumbers);
        Assert.Contains(2, slideNumbers);
    }

    // ── GetTextFormatting: error handling ─────────────────────────────────────────

    [Fact]
    public void GetTextFormatting_SlideNumberOutOfRange_ReturnsFailure()
    {
        var path = CreateMinimalPptx();

        var result = Service.GetTextFormatting(path, slideNumber: 99);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message);
    }

    [Fact]
    public void GetTextFormatting_SlideNumberZero_ReturnsFailure()
    {
        var path = CreateMinimalPptx();

        var result = Service.GetTextFormatting(path, slideNumber: 0);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message);
    }

    [Fact]
    public void GetTextFormatting_FileNotFound_ThrowsException()
    {
        var fakePath = Path.Join(Path.GetTempPath(), "nonexistent-text-formatting.pptx");

        Assert.ThrowsAny<Exception>(() => Service.GetTextFormatting(fakePath));
    }

    [Fact]
    public void GetTextFormatting_EmptyPresentation_ReturnsSuccessWithZeroFormattings()
    {
        var path = CreatePptxWithSlides();

        var result = Service.GetTextFormatting(path);

        Assert.True(result.Success);
        Assert.Equal(0, result.FormattingCount);
        Assert.Contains("no slides", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ── ApplyTextFormatting: font family ──────────────────────────────────────────

    [Fact]
    public void ApplyTextFormatting_FontFamily_AppliedAndReadBack()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        var result = Service.ApplyTextFormatting(path, 1, "Target", fontFamily: "Verdana");

        Assert.True(result.Success);
        Assert.Equal("Apply", result.Action);
        Assert.True(result.FormattingCount > 0);

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal("Verdana", run.FontFamily);
    }

    [Fact]
    public void ApplyTextFormatting_FontFamily_ReplacesExisting()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test", fontFamily: "Arial");

        Service.ApplyTextFormatting(path, 1, "Target", fontFamily: "Calibri");

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal("Calibri", run.FontFamily);
    }

    // ── ApplyTextFormatting: font size ────────────────────────────────────────────

    [Fact]
    public void ApplyTextFormatting_FontSize_PointsConvertedToHundredths()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Service.ApplyTextFormatting(path, 1, "Target", fontSize: 18.0);

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal(18.0, run.FontSize);
    }

    [Fact]
    public void ApplyTextFormatting_FontSize_HalfPointPrecision()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Service.ApplyTextFormatting(path, 1, "Target", fontSize: 10.5);

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal(10.5, run.FontSize);
    }

    // ── ApplyTextFormatting: bold / italic / underline ────────────────────────────

    [Fact]
    public void ApplyTextFormatting_Bold_TurnedOnAndReadBack()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Service.ApplyTextFormatting(path, 1, "Target", bold: true);

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.True(run.Bold);
    }

    [Fact]
    public void ApplyTextFormatting_Italic_TurnedOnAndReadBack()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Service.ApplyTextFormatting(path, 1, "Target", italic: true);

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.True(run.Italic);
    }

    [Fact]
    public void ApplyTextFormatting_Underline_TurnedOnAndReadBack()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Service.ApplyTextFormatting(path, 1, "Target", underline: true);

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.True(run.Underline);
    }

    [Fact]
    public void ApplyTextFormatting_UnderlineOff_SetsToNone()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test", underline: A.TextUnderlineValues.Single);

        Service.ApplyTextFormatting(path, 1, "Target", underline: false);

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.False(run.Underline);
    }

    [Fact]
    public void ApplyTextFormatting_BoldOff_TurnsBoldOff()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test", bold: true);

        Service.ApplyTextFormatting(path, 1, "Target", bold: false);

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.False(run.Bold);
    }

    // ── ApplyTextFormatting: color ────────────────────────────────────────────────

    [Fact]
    public void ApplyTextFormatting_Color_AppliedAsRgbHex()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Service.ApplyTextFormatting(path, 1, "Target", color: "#00FF00");

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal("#00FF00", run.Color);
    }

    [Fact]
    public void ApplyTextFormatting_Color_WithoutHashPrefix()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Service.ApplyTextFormatting(path, 1, "Target", color: "0000FF");

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal("#0000FF", run.Color);
    }

    [Fact]
    public void ApplyTextFormatting_Color_ReplacesExistingColor()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test", colorHex: "FF0000");

        Service.ApplyTextFormatting(path, 1, "Target", color: "#00FF00");

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal("#00FF00", run.Color);
    }

    // ── ApplyTextFormatting: alignment ────────────────────────────────────────────

    [Theory]
    [InlineData("Left", "Left")]
    [InlineData("Center", "Center")]
    [InlineData("Right", "Right")]
    [InlineData("Justify", "Justify")]
    public void ApplyTextFormatting_Alignment_AppliedAndReadBack(string input, string expected)
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Service.ApplyTextFormatting(path, 1, "Target", alignment: input);

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal(expected, run.Alignment);
    }

    [Fact]
    public void ApplyTextFormatting_Alignment_CaseInsensitive()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Service.ApplyTextFormatting(path, 1, "Target", alignment: "center");

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal("Center", run.Alignment);
    }

    [Fact]
    public void ApplyTextFormatting_Alignment_JustifiedAlias()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Service.ApplyTextFormatting(path, 1, "Target", alignment: "justified");

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal("Justify", run.Alignment);
    }

    [Fact]
    public void ApplyTextFormatting_Alignment_InvalidValue_ThrowsArgumentException()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Assert.Throws<ArgumentException>(() =>
            Service.ApplyTextFormatting(path, 1, "Target", alignment: "diagonal"));
    }

    // ── ApplyTextFormatting: no-op / invalid input ────────────────────────────────

    [Fact]
    public void ApplyTextFormatting_NoPropertiesSpecified_ReturnsFailure()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        var result = Service.ApplyTextFormatting(path, 1, "Target");

        Assert.False(result.Success);
        Assert.Contains("No formatting properties specified", result.Message);
    }

    [Fact]
    public void ApplyTextFormatting_NegativeFontSize_ReturnsFailure()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        var result = Service.ApplyTextFormatting(path, 1, "Target", fontSize: -1.0);

        Assert.False(result.Success);
        Assert.Contains("fontSize must be a positive number", result.Message);
    }

    [Fact]
    public void ApplyTextFormatting_ZeroFontSize_ReturnsFailure()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        var result = Service.ApplyTextFormatting(path, 1, "Target", fontSize: 0.0);

        Assert.False(result.Success);
        Assert.Contains("fontSize must be a positive number", result.Message);
    }

    [Fact]
    public void ApplyTextFormatting_InvalidHexColor_ThrowsArgumentException()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Assert.Throws<ArgumentException>(() =>
            Service.ApplyTextFormatting(path, 1, "Target", color: "red"));
    }

    [Fact]
    public void ApplyTextFormatting_InvalidHexColor_TooShort_ThrowsArgumentException()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Assert.Throws<ArgumentException>(() =>
            Service.ApplyTextFormatting(path, 1, "Target", color: "#FFF"));
    }

    [Fact]
    public void ApplyTextFormatting_InvalidHexColor_InvalidCharacters_ThrowsArgumentException()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Assert.Throws<ArgumentException>(() =>
            Service.ApplyTextFormatting(path, 1, "Target", color: "#GGGGGG"));
    }

    [Fact]
    public void ApplyTextFormatting_IdempotentBold_ReturnsZeroModifiedCount()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test", bold: true);

        // Bold is already true — applying bold: true again should report no actual changes
        var result = Service.ApplyTextFormatting(path, 1, "Target", bold: true);

        Assert.True(result.Success);
        Assert.Equal(0, result.FormattingCount);
    }

    [Fact]
    public void ApplyTextFormatting_IdempotentColor_ReturnsZeroModifiedCount()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test", colorHex: "FF0000");

        // Same color — applying #FF0000 again should report no actual changes
        var result = Service.ApplyTextFormatting(path, 1, "Target", color: "#FF0000");

        Assert.True(result.Success);
        Assert.Equal(0, result.FormattingCount);
    }

    // ── GetTextFormatting: scheme / non-RGB colors ────────────────────────────────

    [Fact]
    public void GetTextFormatting_SchemeColor_ReturnedWithSchemePrefix()
    {
        var path = CreateMinimalPptx("Slide 1");

        // Build the fixture and close the document before reading back
        {
            using var doc = PresentationDocument.Open(path, true);
            var slidePart = doc.PresentationPart!.SlideParts.First();
            var tree = slidePart.Slide.CommonSlideData!.ShapeTree!;

            var schemeColor = new A.SchemeColor
            {
                Val = new DocumentFormat.OpenXml.EnumValue<A.SchemeColorValues>(A.SchemeColorValues.Accent1)
            };
            var runProps = new A.RunProperties();
            runProps.InsertAt(new A.SolidFill(schemeColor), 0);

            var paragraph = new A.Paragraph(
                new A.Run(runProps, new A.Text("Themed")),
                new A.EndParagraphRunProperties());
            var textBody = new TextBody(new A.BodyProperties(), new A.ListStyle(), paragraph);

            var nextId = (uint)(tree.Elements<Shape>().Count() + 10);
            tree.Append(new Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = nextId, Name = "SchemeShape" },
                    new P.NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 100000, Y = 100000 },
                        new A.Extents { Cx = 5000000, Cy = 1000000 })),
                textBody));

            slidePart.Slide.Save();
        }

        var result = Service.GetTextFormatting(path, shapeName: "SchemeShape");
        var run = Assert.Single(result.Formattings, f => f.ShapeName == "SchemeShape");
        Assert.NotNull(run.Color);
        Assert.StartsWith("scheme:", run.Color);
    }

    // ── ApplyTextFormatting: multiple properties at once ──────────────────────────

    [Fact]
    public void ApplyTextFormatting_MultipleProperties_AllApplied()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        Service.ApplyTextFormatting(path, 1, "Target",
            fontFamily: "Georgia",
            fontSize: 16.0,
            bold: true,
            italic: true,
            underline: true,
            color: "#336699",
            alignment: "Right");

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal("Georgia", run.FontFamily);
        Assert.Equal(16.0, run.FontSize);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.True(run.Underline);
        Assert.Equal("#336699", run.Color);
        Assert.Equal("Right", run.Alignment);
    }

    // ── ApplyTextFormatting: preserves unset properties ───────────────────────────

    [Fact]
    public void ApplyTextFormatting_PreservesUnsetProperties()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test",
            fontFamily: "Arial", bold: true, colorHex: "FF0000");

        Service.ApplyTextFormatting(path, 1, "Target", italic: true);

        var readBack = Service.GetTextFormatting(path, shapeName: "Target");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "Target");
        Assert.Equal("Arial", run.FontFamily);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.Equal("#FF0000", run.Color);
    }

    // ── ApplyTextFormatting: shape with no existing RunProperties ─────────────────

    [Fact]
    public void ApplyTextFormatting_CreatesRunPropertiesWhenAbsent()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "PlainShape", Paragraphs = ["No formatting"] }
            ]
        });

        var result = Service.ApplyTextFormatting(path, 1, "PlainShape", bold: true, fontSize: 20.0);

        Assert.True(result.Success);
        Assert.True(result.FormattingCount > 0);

        var readBack = Service.GetTextFormatting(path, shapeName: "PlainShape");
        var run = Assert.Single(readBack.Formattings, f => f.ShapeName == "PlainShape");
        Assert.True(run.Bold);
        Assert.Equal(20.0, run.FontSize);
    }

    // ── ApplyTextFormatting: shapes with no text runs ────────────────────────────

    [Fact]
    public void ApplyTextFormatting_ShapeWithNoRuns_ReturnsZeroModified()
    {
        var path = CreateShapeWithNoRuns("EmptyShape");

        var result = Service.ApplyTextFormatting(path, 1, "EmptyShape", bold: true);

        Assert.True(result.Success);
        Assert.Equal(0, result.FormattingCount);
    }

    // ── ApplyTextFormatting: error handling ───────────────────────────────────────

    [Fact]
    public void ApplyTextFormatting_ShapeNotFound_ReturnsFailure()
    {
        var path = CreateMinimalPptx();

        var result = Service.ApplyTextFormatting(path, 1, "NonExistentShape", bold: true);

        Assert.False(result.Success);
        Assert.Contains("No shape named", result.Message);
        Assert.Contains("Available shapes", result.Message);
    }

    [Fact]
    public void ApplyTextFormatting_ShapeNameIsCaseInsensitive()
    {
        var path = CreateFormattedPptx(shapeName: "MyShape", text: "Test");

        var result = Service.ApplyTextFormatting(path, 1, "MYSHAPE", bold: true);

        Assert.True(result.Success);
    }

    [Fact]
    public void ApplyTextFormatting_SlideNotFound_ReturnsFailure()
    {
        var path = CreateMinimalPptx();

        var result = Service.ApplyTextFormatting(path, 99, "AnyShape", bold: true);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message);
    }

    [Fact]
    public void ApplyTextFormatting_SlideNumberZero_ReturnsFailure()
    {
        var path = CreateMinimalPptx();

        var result = Service.ApplyTextFormatting(path, 0, "AnyShape", bold: true);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message);
    }

    [Fact]
    public void ApplyTextFormatting_FileNotFound_ThrowsException()
    {
        var fakePath = Path.Join(Path.GetTempPath(), "nonexistent-apply.pptx");

        Assert.ThrowsAny<Exception>(() =>
            Service.ApplyTextFormatting(fakePath, 1, "Shape", bold: true));
    }

    // ── ApplyTextFormatting: multi-run shapes ────────────────────────────────────

    [Fact]
    public void ApplyTextFormatting_MultipleRuns_AllRunsModified()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TextShapes =
            [
                new TestTextShapeDefinition
                {
                    Name = "MultiRun",
                    ParagraphDefinitions =
                    [
                        new TestParagraphDefinition { Text = "Line 1" },
                        new TestParagraphDefinition { Text = "Line 2" },
                        new TestParagraphDefinition { Text = "Line 3" }
                    ]
                }
            ]
        });

        var result = Service.ApplyTextFormatting(path, 1, "MultiRun", bold: true);

        Assert.True(result.Success);
        Assert.Equal(3, result.FormattingCount);

        var readBack = Service.GetTextFormatting(path, shapeName: "MultiRun");
        Assert.All(readBack.Formattings, f => Assert.True(f.Bold));
    }

    // ── Round-trip: verify OpenXML integrity ─────────────────────────────────────

    [Fact]
    public void ApplyTextFormatting_RoundTrip_OpenXmlDocumentStaysValid()
    {
        var path = CreateFormattedPptx(shapeName: "Target", text: "Test");

        var baselineErrors = CountValidationErrors(path);

        Service.ApplyTextFormatting(path, 1, "Target",
            fontFamily: "Consolas", fontSize: 14.0, bold: true, color: "#112233", alignment: "Center");

        var afterErrors = CountValidationErrors(path);
        Assert.True(afterErrors <= baselineErrors,
            $"Validation errors increased from {baselineErrors} to {afterErrors} after applying formatting.");
    }

    private static int CountValidationErrors(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        return validator.Validate(doc).Count();
    }
}
