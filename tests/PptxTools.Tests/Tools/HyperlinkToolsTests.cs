using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests.Tools;

/// <summary>
/// Tool-level tests for pptx_manage_hyperlinks MCP tool.
/// Written proactively for Issue #114 — hyperlink support.
/// Tests verify JSON output format, error handling, and parameter validation at the MCP tool layer.
/// </summary>
[Trait("Category", "Integration")]
public class HyperlinkToolsTests : PptxTestBase
{
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public HyperlinkToolsTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    // ────────────────────────────────────────────────────────
    // Fixture helpers
    // ────────────────────────────────────────────────────────

    private string CreatePptxWithHyperlink(string shapeName, string displayText, string url, string? tooltip = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"hlink-tool-{Guid.NewGuid()}.pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

        slideLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = SlideLayoutValues.Title
        };
        slideLayoutPart.SlideLayout.CommonSlideData!.Name = "Title Slide";

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new SlideLayoutIdList(
                new SlideLayoutId
                {
                    Id = 2049,
                    RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
                }));
        slideLayoutPart.AddPart(slideMasterPart);

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.AddPart(slideLayoutPart);

        // Create shape with hyperlink
        var run = new A.Run(
            new A.RunProperties { Language = "en-US" },
            new A.Text(displayText));

        var relationship = slidePart.AddHyperlinkRelationship(new Uri(url), true);
        var hlinkClick = new A.HyperlinkOnClick { Id = relationship.Id };
        if (tooltip is not null)
            hlinkClick.Tooltip = tooltip;
        run.RunProperties!.Append(hlinkClick);

        var shapeTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()),
            new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 2, Name = shapeName },
                    new P.NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 457200, Y = 457200 },
                        new A.Extents { Cx = 8229600, Cy = 685800 }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(run))));

        slidePart.Slide = new Slide(
            new CommonSlideData(shapeTree),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        var slideIdList = new SlideIdList(
            new SlideId { Id = 256, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

        presentationPart.Presentation = new Presentation(
            slideIdList,
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });

        var slideMasterIdList = new SlideMasterIdList(
            new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });
        presentationPart.Presentation.InsertAt(slideMasterIdList, 0);
        presentationPart.Presentation.Save();

        return path;
    }

    private string CreatePptxWithNamedShape(string shapeName)
    {
        return CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Test",
            TextShapes = [new TestTextShapeDefinition { Name = shapeName, Paragraphs = ["Sample text"] }]
        });
    }

    // ────────────────────────────────────────────────────────
    // Get action: returns structured JSON
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task ManageHyperlinks_Get_ReturnsStructuredResult()
    {
        var path = CreatePptxWithHyperlink("Link Shape", "Visit", "https://github.com");

        var result = await _tools.pptx_manage_hyperlinks(path, HyperlinkAction.Get, slideNumber: 1);

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("Get", parsed.Action);
        Assert.True(parsed.HyperlinkCount >= 1);
        Assert.NotNull(parsed.Hyperlinks);
        Assert.Contains(parsed.Hyperlinks, h => h.Url == "https://github.com");
    }

    [Fact]
    public async Task ManageHyperlinks_Get_EmptyPresentation_ReturnsZeroCount()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_manage_hyperlinks(path, HyperlinkAction.Get, slideNumber: 1);

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal(0, parsed.HyperlinkCount);
    }

    [Fact]
    public async Task ManageHyperlinks_Get_NoSlideFilter_ReturnsAll()
    {
        var path = CreatePptxWithHyperlink("Link Shape", "Click", "https://example.com");

        var result = await _tools.pptx_manage_hyperlinks(path, HyperlinkAction.Get);

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.True(parsed.HyperlinkCount >= 1);
    }

    // ────────────────────────────────────────────────────────
    // Add action: returns structured success
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task ManageHyperlinks_Add_ReturnsStructuredSuccess()
    {
        var path = CreatePptxWithNamedShape("Target Shape");

        var result = await _tools.pptx_manage_hyperlinks(
            path, HyperlinkAction.Add,
            slideNumber: 1,
            shapeName: "Target Shape",
            url: "https://github.com");

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("Add", parsed.Action);
        Assert.Equal("https://github.com", parsed.Url);
    }

    [Fact]
    public async Task ManageHyperlinks_Add_IsVerifiableViaGet()
    {
        var path = CreatePptxWithNamedShape("Target Shape");

        await _tools.pptx_manage_hyperlinks(
            path, HyperlinkAction.Add,
            slideNumber: 1,
            shapeName: "Target Shape",
            url: "https://github.com");

        var getResult = await _tools.pptx_manage_hyperlinks(path, HyperlinkAction.Get, slideNumber: 1);
        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(getResult);
        Assert.NotNull(parsed);
        Assert.True(parsed.HyperlinkCount >= 1);
        Assert.Contains(parsed.Hyperlinks!, h => h.Url == "https://github.com");
    }

    [Fact]
    public async Task ManageHyperlinks_Add_WithTooltip_Succeeds()
    {
        var path = CreatePptxWithNamedShape("Tip Shape");

        var result = await _tools.pptx_manage_hyperlinks(
            path, HyperlinkAction.Add,
            slideNumber: 1,
            shapeName: "Tip Shape",
            url: "https://example.com",
            tooltip: "Helpful tip");

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
    }

    // ────────────────────────────────────────────────────────
    // Update action: returns structured success
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task ManageHyperlinks_Update_ReturnsStructuredSuccess()
    {
        var path = CreatePptxWithHyperlink("Link Shape", "Click", "https://old.com");

        var result = await _tools.pptx_manage_hyperlinks(
            path, HyperlinkAction.Update,
            slideNumber: 1,
            shapeName: "Link Shape",
            url: "https://new.com");

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("Update", parsed.Action);
        Assert.Equal("https://new.com", parsed.Url);
    }

    [Fact]
    public async Task ManageHyperlinks_Update_VerifiesUrlChanged()
    {
        var path = CreatePptxWithHyperlink("Link Shape", "Click", "https://old.com");

        await _tools.pptx_manage_hyperlinks(
            path, HyperlinkAction.Update,
            slideNumber: 1,
            shapeName: "Link Shape",
            url: "https://new.com");

        var links = Service.GetHyperlinks(path);
        var link = Assert.Single(links);
        Assert.Equal("https://new.com", link.Url);
    }

    // ────────────────────────────────────────────────────────
    // Remove action: returns structured success
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task ManageHyperlinks_Remove_ReturnsStructuredSuccess()
    {
        var path = CreatePptxWithHyperlink("Link Shape", "Click", "https://example.com");

        var result = await _tools.pptx_manage_hyperlinks(
            path, HyperlinkAction.Remove,
            slideNumber: 1,
            shapeName: "Link Shape");

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("Remove", parsed.Action);
    }

    [Fact]
    public async Task ManageHyperlinks_Remove_VerifiesHyperlinkGone()
    {
        var path = CreatePptxWithHyperlink("Link Shape", "Click", "https://example.com");

        await _tools.pptx_manage_hyperlinks(
            path, HyperlinkAction.Remove,
            slideNumber: 1,
            shapeName: "Link Shape");

        var links = Service.GetHyperlinks(path);
        Assert.Empty(links);
    }

    // ────────────────────────────────────────────────────────
    // Error handling: file not found
    // ────────────────────────────────────────────────────────

    [Theory]
    [InlineData(HyperlinkAction.Get)]
    [InlineData(HyperlinkAction.Add)]
    [InlineData(HyperlinkAction.Update)]
    [InlineData(HyperlinkAction.Remove)]
    public async Task ManageHyperlinks_FileNotFound_ReturnsStructuredError(HyperlinkAction action)
    {
        var fakePath = Path.Combine(Path.GetTempPath(), "nonexistent-hyperlink-test.pptx");

        var result = await _tools.pptx_manage_hyperlinks(
            fakePath, action,
            slideNumber: 1,
            shapeName: "Shape",
            url: "https://example.com");

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("not found", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────
    // Error handling: missing required parameters
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task ManageHyperlinks_Add_MissingSlideNumber_ReturnsStructuredError()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_manage_hyperlinks(
            path, HyperlinkAction.Add,
            slideNumber: null,
            shapeName: "Shape",
            url: "https://example.com");

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("slideNumber", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ManageHyperlinks_Add_MissingShapeName_ReturnsStructuredError()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_manage_hyperlinks(
            path, HyperlinkAction.Add,
            slideNumber: 1,
            shapeName: null,
            url: "https://example.com");

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("shapeName", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ManageHyperlinks_Add_MissingUrl_ReturnsStructuredError()
    {
        var path = CreatePptxWithNamedShape("My Shape");

        var result = await _tools.pptx_manage_hyperlinks(
            path, HyperlinkAction.Add,
            slideNumber: 1,
            shapeName: "My Shape",
            url: null);

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("url", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ManageHyperlinks_Remove_MissingShapeName_ReturnsStructuredError()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_manage_hyperlinks(
            path, HyperlinkAction.Remove,
            slideNumber: 1,
            shapeName: null);

        var parsed = JsonSerializer.Deserialize<HyperlinkResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("shapeName", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }
}
