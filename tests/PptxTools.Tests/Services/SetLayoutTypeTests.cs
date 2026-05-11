using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests.Services;

/// <summary>
/// Service-level tests for SetLayoutType (pptx_manage_layouts SetType action).
/// </summary>
[Trait("Category", "Unit")]
public class SetLayoutTypeTests : PptxTestBase
{
    // ──────────────────────────────────────────────────────────
    //  1. Happy-path: set a known layout type
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void SetLayoutType_KnownType_ReturnsSuccess()
    {
        var path = CreatePptxWithNamedLayouts("Title Slide", "Section Header", "Content");

        var result = Service.SetLayoutType(path, "Title Slide", "title");

        Assert.True(result.Success);
    }

    [Fact]
    public void SetLayoutType_KnownType_SetsTypeOnDisk()
    {
        var path = CreatePptxWithNamedLayouts("Title Slide", "Section Header", "Content");

        Service.SetLayoutType(path, "Title Slide", "title");

        // Verify the type was persisted by re-opening the file.
        using var doc = PresentationDocument.Open(path, false);
        var layout = doc.PresentationPart!.SlideMasterParts
            .SelectMany(m => m.SlideLayoutParts)
            .First(lp => lp.SlideLayout.CommonSlideData?.Name?.Value == "Title Slide");

        Assert.Equal(SlideLayoutValues.Title, layout.SlideLayout.Type?.Value);
    }

    [Fact]
    public void SetLayoutType_SectionHeader_SetsCorrectValue()
    {
        var path = CreatePptxWithNamedLayouts("Title Slide", "Section Header", "Content");

        var result = Service.SetLayoutType(path, "Section Header", "secHead");

        Assert.True(result.Success);

        using var doc = PresentationDocument.Open(path, false);
        var layout = doc.PresentationPart!.SlideMasterParts
            .SelectMany(m => m.SlideLayoutParts)
            .First(lp => lp.SlideLayout.CommonSlideData?.Name?.Value == "Section Header");

        Assert.Equal(SlideLayoutValues.SectionHeader, layout.SlideLayout.Type?.Value);
    }

    [Fact]
    public void SetLayoutType_Blank_SetsCorrectValue()
    {
        var path = CreatePptxWithNamedLayouts("Title Slide", "Section Header", "Blank");

        Service.SetLayoutType(path, "Blank", "blank");

        using var doc = PresentationDocument.Open(path, false);
        var layout = doc.PresentationPart!.SlideMasterParts
            .SelectMany(m => m.SlideLayoutParts)
            .First(lp => lp.SlideLayout.CommonSlideData?.Name?.Value == "Blank");

        Assert.Equal(SlideLayoutValues.Blank, layout.SlideLayout.Type?.Value);
    }

    [Fact]
    public void SetLayoutType_TwoObj_SetsCorrectValue()
    {
        var path = CreatePptxWithNamedLayouts("Two Content");

        Service.SetLayoutType(path, "Two Content", "twoObj");

        using var doc = PresentationDocument.Open(path, false);
        var layout = doc.PresentationPart!.SlideMasterParts
            .SelectMany(m => m.SlideLayoutParts)
            .First(lp => lp.SlideLayout.CommonSlideData?.Name?.Value == "Two Content");

        Assert.Equal(SlideLayoutValues.TwoObjects, layout.SlideLayout.Type?.Value);
    }

    // ──────────────────────────────────────────────────────────
    //  2. Result fields are populated correctly
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void SetLayoutType_Result_ContainsLayoutNameAndType()
    {
        var path = CreatePptxWithNamedLayouts("My Layout");

        var result = Service.SetLayoutType(path, "My Layout", "obj");

        Assert.Equal("My Layout", result.LayoutName);
        Assert.Equal("obj", result.LayoutType);
        Assert.Equal(path, result.FilePath);
    }

    // ──────────────────────────────────────────────────────────
    //  3. Case-insensitive layout name lookup
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void SetLayoutType_LayoutNameCaseInsensitive_Succeeds()
    {
        var path = CreatePptxWithNamedLayouts("Title Slide");

        var result = Service.SetLayoutType(path, "title slide", "title");

        Assert.True(result.Success);
    }

    // ──────────────────────────────────────────────────────────
    //  4. Case-insensitive type value lookup
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void SetLayoutType_TypeValueCaseInsensitive_Succeeds()
    {
        var path = CreatePptxWithNamedLayouts("Title Slide");

        var result = Service.SetLayoutType(path, "Title Slide", "TITLE");

        Assert.True(result.Success);
    }

    // ──────────────────────────────────────────────────────────
    //  5. Layout not found returns failure
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void SetLayoutType_LayoutNotFound_ReturnsFailure()
    {
        var path = CreatePptxWithNamedLayouts("Title Slide");

        var result = Service.SetLayoutType(path, "Nonexistent Layout", "title");

        Assert.False(result.Success);
        Assert.Contains("not found", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ──────────────────────────────────────────────────────────
    //  6. Invalid type value returns failure with hint
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void SetLayoutType_InvalidType_ReturnsFailure()
    {
        var path = CreatePptxWithNamedLayouts("Title Slide");

        var result = Service.SetLayoutType(path, "Title Slide", "invalidType");

        Assert.False(result.Success);
        Assert.Contains("invalidType", result.Message);
    }

    [Fact]
    public void SetLayoutType_InvalidType_MessageContainsValidValues()
    {
        var path = CreatePptxWithNamedLayouts("Title Slide");

        var result = Service.SetLayoutType(path, "Title Slide", "badvalue");

        // The error message should list at least some valid values.
        Assert.Contains("title", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ──────────────────────────────────────────────────────────
    //  7. Overwrite existing type value
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void SetLayoutType_OverwriteExistingType_Succeeds()
    {
        // CreatePptxWithNamedLayouts sets "Title Slide" layout to SlideLayoutValues.Title.
        var path = CreatePptxWithNamedLayouts("Title Slide");

        // Change it to "secHead".
        var result = Service.SetLayoutType(path, "Title Slide", "secHead");

        Assert.True(result.Success);

        using var doc = PresentationDocument.Open(path, false);
        var layout = doc.PresentationPart!.SlideMasterParts
            .SelectMany(m => m.SlideLayoutParts)
            .First(lp => lp.SlideLayout.CommonSlideData?.Name?.Value == "Title Slide");

        Assert.Equal(SlideLayoutValues.SectionHeader, layout.SlideLayout.Type?.Value);
    }

    // ──────────────────────────────────────────────────────────
    //  8. Set type on layout without an existing type attribute
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void SetLayoutType_NoExistingType_SetsType()
    {
        var path = CreatePptxWithUntypedLayout("My Content Layout");

        var result = Service.SetLayoutType(path, "My Content Layout", "obj");

        Assert.True(result.Success);

        using var doc = PresentationDocument.Open(path, false);
        var layout = doc.PresentationPart!.SlideMasterParts
            .SelectMany(m => m.SlideLayoutParts)
            .First(lp => lp.SlideLayout.CommonSlideData?.Name?.Value == "My Content Layout");

        Assert.Equal(SlideLayoutValues.Object, layout.SlideLayout.Type?.Value);
    }

    // ──────────────────────────────────────────────────────────
    //  Helpers
    // ──────────────────────────────────────────────────────────

    /// <summary>
    /// Creates a minimal PPTX with one layout per name, each starting without a type.
    /// </summary>
    private string CreatePptxWithNamedLayouts(params string[] names)
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();

        var layoutIds = new List<SlideLayoutId>();
        uint idCounter = 2049;
        foreach (var name in names)
        {
            var layoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            layoutPart.SlideLayout = new SlideLayout(
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
            layoutPart.SlideLayout.CommonSlideData!.Name = name;
            layoutPart.AddPart(slideMasterPart);
            layoutIds.Add(new SlideLayoutId { Id = idCounter++, RelationshipId = slideMasterPart.GetIdOfPart(layoutPart) });
        }

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
            new SlideLayoutIdList(layoutIds.ToArray()));

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        // Wire the slide to the first layout so the package is valid.
        var firstLayoutPart = slideMasterPart.SlideLayoutParts.First();
        slidePart.AddPart(firstLayoutPart);
        slidePart.Slide = new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))));

        var slideMasterIdList = new SlideMasterIdList(
            new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });
        var slideIdList = new SlideIdList(
            new SlideId { Id = 256, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

        presentationPart.Presentation = new Presentation(
            slideIdList,
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });
        presentationPart.Presentation.InsertAt(slideMasterIdList, 0);
        presentationPart.Presentation.Save();

        return path;
    }

    /// <summary>
    /// Creates a minimal PPTX where the named layout has no <c>type</c> attribute set.
    /// </summary>
    private string CreatePptxWithUntypedLayout(string layoutName)
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();

        var layoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        // Explicitly omit the Type property to simulate layouts that have no type attribute.
        layoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()));
        layoutPart.SlideLayout.CommonSlideData!.Name = layoutName;
        layoutPart.AddPart(slideMasterPart);

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
                    RelationshipId = slideMasterPart.GetIdOfPart(layoutPart)
                }));

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.AddPart(layoutPart);
        slidePart.Slide = new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))));

        var slideMasterIdList = new SlideMasterIdList(
            new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });
        var slideIdList = new SlideIdList(
            new SlideId { Id = 256, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

        presentationPart.Presentation = new Presentation(
            slideIdList,
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });
        presentationPart.Presentation.InsertAt(slideMasterIdList, 0);
        presentationPart.Presentation.Save();

        return path;
    }
}
