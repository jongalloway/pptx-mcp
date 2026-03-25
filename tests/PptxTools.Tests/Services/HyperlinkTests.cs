using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests.Services;

/// <summary>
/// Service-level tests for hyperlink operations: GetHyperlinks, AddHyperlink, UpdateHyperlink, RemoveHyperlink.
/// Written proactively for Issue #114 — hyperlink support.
/// </summary>
[Trait("Category", "Unit")]
public class HyperlinkTests : PptxTestBase
{
    // ────────────────────────────────────────────────────────
    // Fixture helpers
    // ────────────────────────────────────────────────────────

    /// <summary>Creates a presentation with shapes that have hyperlinks attached at the run level.</summary>
    private string CreatePptxWithHyperlinks(params HyperlinkFixture[] fixtures)
    {
        var path = Path.Combine(Path.GetTempPath(), $"hyperlink-test-{Guid.NewGuid()}.pptx");
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

        // Group fixtures by slide number
        var slideGroups = fixtures.GroupBy(f => f.SlideNumber).OrderBy(g => g.Key).ToList();
        if (slideGroups.Count == 0)
            slideGroups = [new FakeGrouping(1, [])];

        var slideIdList = new SlideIdList();
        uint nextSlideId = 256;

        foreach (var group in slideGroups)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(slideLayoutPart);

            var shapeTree = new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()));

            uint shapeId = 2;
            foreach (var fixture in group)
            {
                var shape = CreateTextShapeWithHyperlink(slidePart, shapeId++, fixture);
                shapeTree.Append(shape);
            }

            slidePart.Slide = new Slide(
                new CommonSlideData(shapeTree),
                new P.ColorMapOverride(new A.MasterColorMapping()));

            slideIdList.Append(new SlideId
            {
                Id = nextSlideId++,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

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

    /// <summary>Creates a presentation with named text shapes but no hyperlinks.</summary>
    private string CreatePptxWithNamedShapes(params string[] shapeNames)
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Test Slide",
            TextShapes = shapeNames.Select(name => new TestTextShapeDefinition
            {
                Name = name,
                Paragraphs = [$"Content of {name}"]
            }).ToList()
        });
        return path;
    }

    private static P.Shape CreateTextShapeWithHyperlink(SlidePart slidePart, uint shapeId, HyperlinkFixture fixture)
    {
        var run = new A.Run(
            new A.RunProperties { Language = "en-US" },
            new A.Text(fixture.DisplayText));

        if (!string.IsNullOrEmpty(fixture.Url))
        {
            var relationship = slidePart.AddHyperlinkRelationship(new Uri(fixture.Url), true);
            var hlinkClick = new A.HyperlinkOnClick { Id = relationship.Id };
            if (!string.IsNullOrEmpty(fixture.Tooltip))
                hlinkClick.Tooltip = fixture.Tooltip;
            run.RunProperties!.Append(hlinkClick);
        }

        var paragraph = new A.Paragraph(run);

        var textBody = new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            paragraph);

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = fixture.ShapeName },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 457200, Y = 457200 },
                    new A.Extents { Cx = 8229600, Cy = 685800 }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            textBody);
    }

    // ────────────────────────────────────────────────────────
    // GetHyperlinks: empty / no hyperlinks
    // ────────────────────────────────────────────────────────

    [Fact]
    public void GetHyperlinks_EmptyPresentation_ReturnsEmptyList()
    {
        var path = CreateMinimalPptx();

        var result = Service.GetHyperlinks(path);

        Assert.NotNull(result);
        Assert.Empty(result);
    }

    [Fact]
    public void GetHyperlinks_ShapesWithoutHyperlinks_ReturnsEmptyList()
    {
        var path = CreatePptxWithNamedShapes("Shape A", "Shape B");

        var result = Service.GetHyperlinks(path);

        Assert.NotNull(result);
        Assert.Empty(result);
    }

    // ────────────────────────────────────────────────────────
    // GetHyperlinks: external URL hyperlinks
    // ────────────────────────────────────────────────────────

    [Fact]
    public void GetHyperlinks_ExternalUrl_ReturnsHyperlinkInfo()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Link Shape", "Visit GitHub", "https://github.com"));

        var result = Service.GetHyperlinks(path);

        var link = Assert.Single(result);
        Assert.Equal(1, link.SlideNumber);
        Assert.Equal("Link Shape", link.ShapeName);
        Assert.Equal("https://github.com", link.Url);
    }

    [Fact]
    public void GetHyperlinks_ExternalUrl_IncludesDisplayText()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Link Shape", "Click Here", "https://example.com"));

        var result = Service.GetHyperlinks(path);

        var link = Assert.Single(result);
        Assert.Equal("Click Here", link.Text);
    }

    [Fact]
    public void GetHyperlinks_ExternalUrl_IncludesShapeName()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "My Custom Shape", "Link Text", "https://example.com"));

        var result = Service.GetHyperlinks(path);

        var link = Assert.Single(result);
        Assert.Equal("My Custom Shape", link.ShapeName);
    }

    // ────────────────────────────────────────────────────────
    // GetHyperlinks: mailto hyperlinks
    // ────────────────────────────────────────────────────────

    [Fact]
    public void GetHyperlinks_MailtoLink_ReturnsHyperlinkInfo()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Email Link", "Contact Us", "mailto:support@example.com"));

        var result = Service.GetHyperlinks(path);

        var link = Assert.Single(result);
        Assert.Equal("mailto:support@example.com", link.Url);
    }

    // ────────────────────────────────────────────────────────
    // GetHyperlinks: slide number filtering
    // ────────────────────────────────────────────────────────

    [Fact]
    public void GetHyperlinks_FilterBySlideNumber_ReturnsOnlyMatchingSlide()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Slide1 Link", "Link 1", "https://slide1.com"),
            new HyperlinkFixture(2, "Slide2 Link", "Link 2", "https://slide2.com"));

        var result = Service.GetHyperlinks(path, slideNumber: 1);

        var link = Assert.Single(result);
        Assert.Equal("https://slide1.com", link.Url);
        Assert.Equal(1, link.SlideNumber);
    }

    [Fact]
    public void GetHyperlinks_NoFilter_ReturnsAllHyperlinksAcrossSlides()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Slide1 Link", "Link 1", "https://slide1.com"),
            new HyperlinkFixture(2, "Slide2 Link", "Link 2", "https://slide2.com"));

        var result = Service.GetHyperlinks(path);

        Assert.Equal(2, result.Count);
        Assert.Contains(result, l => l.Url == "https://slide1.com");
        Assert.Contains(result, l => l.Url == "https://slide2.com");
    }

    [Fact]
    public void GetHyperlinks_FilterBySlideNumber_NonexistentSlide_ReturnsEmpty()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Link Shape", "Click", "https://example.com"));

        var result = Service.GetHyperlinks(path, slideNumber: 99);

        Assert.Empty(result);
    }

    // ────────────────────────────────────────────────────────
    // GetHyperlinks: tooltip
    // ────────────────────────────────────────────────────────

    [Fact]
    public void GetHyperlinks_WithTooltip_IncludesTooltipInResult()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Tip Shape", "Hover Me", "https://example.com", "Click to visit"));

        var result = Service.GetHyperlinks(path);

        var link = Assert.Single(result);
        Assert.Equal("Click to visit", link.Tooltip);
    }

    // ────────────────────────────────────────────────────────
    // GetHyperlinks: multiple hyperlinks on one slide
    // ────────────────────────────────────────────────────────

    [Fact]
    public void GetHyperlinks_MultipleOnSameSlide_ReturnsAll()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Shape A", "GitHub", "https://github.com"),
            new HyperlinkFixture(1, "Shape B", "Google", "https://google.com"),
            new HyperlinkFixture(1, "Shape C", "Email", "mailto:test@example.com"));

        var result = Service.GetHyperlinks(path);

        Assert.Equal(3, result.Count);
    }

    // ────────────────────────────────────────────────────────
    // GetHyperlinks: error handling
    // ────────────────────────────────────────────────────────

    [Fact]
    public void GetHyperlinks_FileNotFound_Throws()
    {
        Assert.ThrowsAny<Exception>(() => Service.GetHyperlinks("C:\\nonexistent\\file.pptx"));
    }

    // ────────────────────────────────────────────────────────
    // AddHyperlink: happy path
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AddHyperlink_ToExistingShape_ReturnsSuccess()
    {
        var path = CreatePptxWithNamedShapes("Target Shape");

        var result = Service.AddHyperlink(path, 1, "Target Shape", "https://github.com");

        Assert.True(result.Success);
        Assert.Equal("Add", result.Action);
        Assert.Equal(1, result.SlideNumber);
        Assert.Equal("Target Shape", result.ShapeName);
        Assert.Equal("https://github.com", result.Url);
    }

    [Fact]
    public void AddHyperlink_IsDiscoverableViaGetHyperlinks()
    {
        var path = CreatePptxWithNamedShapes("Clickable");

        // Before: no hyperlinks
        Assert.Empty(Service.GetHyperlinks(path));

        Service.AddHyperlink(path, 1, "Clickable", "https://example.com");

        // After: one hyperlink
        var links = Service.GetHyperlinks(path);
        Assert.Single(links);
        Assert.Equal("https://example.com", links[0].Url);
    }

    [Fact]
    public void AddHyperlink_WithTooltip_StoresToolTip()
    {
        var path = CreatePptxWithNamedShapes("Tip Shape");

        var result = Service.AddHyperlink(path, 1, "Tip Shape", "https://example.com", tooltip: "Visit site");

        Assert.True(result.Success);

        var links = Service.GetHyperlinks(path);
        var link = Assert.Single(links);
        Assert.Equal("Visit site", link.Tooltip);
    }

    [Fact]
    public void AddHyperlink_MailtoUrl_Succeeds()
    {
        var path = CreatePptxWithNamedShapes("Email Shape");

        var result = Service.AddHyperlink(path, 1, "Email Shape", "mailto:contact@example.com");

        Assert.True(result.Success);

        var links = Service.GetHyperlinks(path);
        var link = Assert.Single(links);
        Assert.Equal("mailto:contact@example.com", link.Url);
    }

    // ────────────────────────────────────────────────────────
    // AddHyperlink: error handling
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AddHyperlink_NonExistentShape_Throws()
    {
        var path = CreatePptxWithNamedShapes("Real Shape");

        Assert.Throws<ArgumentException>(() =>
            Service.AddHyperlink(path, 1, "Ghost Shape", "https://example.com"));
    }

    [Fact]
    public void AddHyperlink_InvalidSlideNumber_ThrowsOutOfRange()
    {
        var path = CreatePptxWithNamedShapes("Shape A");

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            Service.AddHyperlink(path, 99, "Shape A", "https://example.com"));
    }

    [Fact]
    public void AddHyperlink_SlideNumberZero_ThrowsOutOfRange()
    {
        var path = CreatePptxWithNamedShapes("Shape A");

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            Service.AddHyperlink(path, 0, "Shape A", "https://example.com"));
    }

    [Fact]
    public void AddHyperlink_FileNotFound_Throws()
    {
        Assert.ThrowsAny<Exception>(() =>
            Service.AddHyperlink("C:\\nonexistent\\file.pptx", 1, "Shape", "https://example.com"));
    }

    // ────────────────────────────────────────────────────────
    // UpdateHyperlink: happy path
    // ────────────────────────────────────────────────────────

    [Fact]
    public void UpdateHyperlink_ChangesUrl()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Link Shape", "Click", "https://old-url.com"));

        var result = Service.UpdateHyperlink(path, 1, "Link Shape", "https://new-url.com");

        Assert.True(result.Success);
        Assert.Equal("Update", result.Action);
        Assert.Equal("https://new-url.com", result.Url);

        var links = Service.GetHyperlinks(path);
        var link = Assert.Single(links);
        Assert.Equal("https://new-url.com", link.Url);
    }

    [Fact]
    public void UpdateHyperlink_ChangesToolTip()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Link Shape", "Click", "https://example.com", "Old Tip"));

        var result = Service.UpdateHyperlink(path, 1, "Link Shape", "https://example.com", newTooltip: "New Tip");

        Assert.True(result.Success);

        var links = Service.GetHyperlinks(path);
        var link = Assert.Single(links);
        Assert.Equal("New Tip", link.Tooltip);
    }

    [Fact]
    public void UpdateHyperlink_PreservesOtherHyperlinks()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Shape A", "Link A", "https://a.com"),
            new HyperlinkFixture(1, "Shape B", "Link B", "https://b.com"));

        Service.UpdateHyperlink(path, 1, "Shape A", "https://updated-a.com");

        var links = Service.GetHyperlinks(path);
        Assert.Equal(2, links.Count);
        Assert.Contains(links, l => l.Url == "https://updated-a.com" && l.ShapeName == "Shape A");
        Assert.Contains(links, l => l.Url == "https://b.com" && l.ShapeName == "Shape B");
    }

    [Fact]
    public void UpdateHyperlink_ShapeWithNoHyperlink_ThrowsInvalidOperation()
    {
        var path = CreatePptxWithNamedShapes("Plain Shape");

        Assert.Throws<InvalidOperationException>(() =>
            Service.UpdateHyperlink(path, 1, "Plain Shape", "https://example.com"));
    }

    [Fact]
    public void UpdateHyperlink_NonExistentShape_ThrowsArgumentException()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Link Shape", "Click", "https://example.com"));

        Assert.Throws<ArgumentException>(() =>
            Service.UpdateHyperlink(path, 1, "Ghost Shape", "https://new-url.com"));
    }

    // ────────────────────────────────────────────────────────
    // RemoveHyperlink: happy path
    // ────────────────────────────────────────────────────────

    [Fact]
    public void RemoveHyperlink_RemovesFromShape_ReturnsSuccess()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Link Shape", "Click", "https://example.com"));

        var result = Service.RemoveHyperlink(path, 1, "Link Shape");

        Assert.True(result.Success);
        Assert.Equal("Remove", result.Action);
        Assert.True(result.HyperlinkCount >= 1);

        var links = Service.GetHyperlinks(path);
        Assert.Empty(links);
    }

    [Fact]
    public void RemoveHyperlink_NoLongerAppearsInGetHyperlinks()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Shape A", "Link A", "https://a.com"),
            new HyperlinkFixture(1, "Shape B", "Link B", "https://b.com"));

        Service.RemoveHyperlink(path, 1, "Shape A");

        var links = Service.GetHyperlinks(path);
        var remaining = Assert.Single(links);
        Assert.Equal("Shape B", remaining.ShapeName);
        Assert.Equal("https://b.com", remaining.Url);
    }

    [Fact]
    public void RemoveHyperlink_PreservesShapeText()
    {
        var path = CreatePptxWithHyperlinks(
            new HyperlinkFixture(1, "Link Shape", "Important Text", "https://example.com"));

        Service.RemoveHyperlink(path, 1, "Link Shape");

        // Shape should still exist with its text, just no hyperlink
        var slideContent = Service.GetSlideContent(path, 0);
        var shape = slideContent.Shapes.FirstOrDefault(s => s.Name == "Link Shape");
        Assert.NotNull(shape);
        Assert.Contains("Important Text", shape.Text);
    }

    [Fact]
    public void RemoveHyperlink_ShapeWithNoHyperlink_ThrowsInvalidOperation()
    {
        var path = CreatePptxWithNamedShapes("Plain Shape");

        Assert.Throws<InvalidOperationException>(() =>
            Service.RemoveHyperlink(path, 1, "Plain Shape"));
    }

    [Fact]
    public void RemoveHyperlink_NonExistentShape_ThrowsArgumentException()
    {
        var path = CreatePptxWithNamedShapes("Real Shape");

        Assert.Throws<ArgumentException>(() =>
            Service.RemoveHyperlink(path, 1, "Ghost Shape"));
    }

    // ────────────────────────────────────────────────────────
    // Round-trip: Add → Get → Update → Get → Remove → Get
    // ────────────────────────────────────────────────────────

    [Fact]
    public void Hyperlink_FullRoundTrip_AddUpdateRemove()
    {
        var path = CreatePptxWithNamedShapes("Round Trip Shape");

        // Add
        var addResult = Service.AddHyperlink(path, 1, "Round Trip Shape", "https://initial.com", tooltip: "Initial");
        Assert.True(addResult.Success);
        var afterAdd = Service.GetHyperlinks(path);
        var link = Assert.Single(afterAdd);
        Assert.Equal("https://initial.com", link.Url);
        Assert.Equal("Initial", link.Tooltip);

        // Update
        var updateResult = Service.UpdateHyperlink(path, 1, "Round Trip Shape", "https://updated.com", newTooltip: "Updated");
        Assert.True(updateResult.Success);
        var afterUpdate = Service.GetHyperlinks(path);
        link = Assert.Single(afterUpdate);
        Assert.Equal("https://updated.com", link.Url);
        Assert.Equal("Updated", link.Tooltip);

        // Remove
        var removeResult = Service.RemoveHyperlink(path, 1, "Round Trip Shape");
        Assert.True(removeResult.Success);
        var afterRemove = Service.GetHyperlinks(path);
        Assert.Empty(afterRemove);
    }

    // ────────────────────────────────────────────────────────
    // Fixture types
    // ────────────────────────────────────────────────────────

    private sealed record HyperlinkFixture(
        int SlideNumber,
        string ShapeName,
        string DisplayText,
        string Url,
        string? Tooltip = null);

    /// <summary>Minimal IGrouping implementation for the fixture helper.</summary>
    private sealed class FakeGrouping(int key, IEnumerable<HyperlinkFixture> items) : IGrouping<int, HyperlinkFixture>
    {
        public int Key => key;
        public IEnumerator<HyperlinkFixture> GetEnumerator() => items.GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
