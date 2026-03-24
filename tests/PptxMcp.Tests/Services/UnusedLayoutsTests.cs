using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Tests.Services;

[Trait("Category", "Unit")]
public class UnusedLayoutsTests : PptxTestBase
{
    // ──────────────────────────────────────────────────────────
    //  1. Minimal PPTX — 1 master, 1 layout, all used
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_MinimalPptx_OneMasterOneLayout()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(1, result.TotalMasters);
        Assert.Equal(1, result.TotalLayouts);
    }

    [Fact]
    public void FindUnusedLayouts_MinimalPptx_AllUsed()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(0, result.UnusedMasterCount);
        Assert.Equal(0, result.UnusedLayoutCount);
    }

    // ──────────────────────────────────────────────────────────
    //  2. Success always true for valid PPTX files
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_ValidFile_SuccessIsTrue()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.True(result.Success);
    }

    [Fact]
    public void FindUnusedLayouts_MultiSlideFile_SuccessIsTrue()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" },
            new TestSlideDefinition { TitleText = "Slide 3" });

        var result = Service.FindUnusedLayouts(path);

        Assert.True(result.Success);
    }

    // ──────────────────────────────────────────────────────────
    //  3. Master is used when any slide references its layout
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_SlideReferencesLayout_MasterIsUsed()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        var master = Assert.Single(result.Masters);
        Assert.True(master.IsUsed);
    }

    [Fact]
    public void FindUnusedLayouts_SlideReferencesLayout_MasterUsedLayoutCountIsOne()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        var master = Assert.Single(result.Masters);
        Assert.Equal(1, master.UsedLayoutCount);
    }

    // ──────────────────────────────────────────────────────────
    //  4. Layout referenced by slide — ReferencedBySlides contains slide number
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_SingleSlide_LayoutReferencedBySlideOne()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        var layout = Assert.Single(result.Layouts);
        Assert.True(layout.IsUsed);
        Assert.Contains(1, layout.ReferencedBySlides);
    }

    // ──────────────────────────────────────────────────────────
    //  5. Multi-slide — all slides using same layout
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_ThreeSlidesSameLayout_ReferencedBySlidesContainsAll()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" },
            new TestSlideDefinition { TitleText = "C" });

        var result = Service.FindUnusedLayouts(path);

        var usedLayout = Assert.Single(result.Layouts, l => l.IsUsed);
        Assert.Equal(new[] { 1, 2, 3 }, usedLayout.ReferencedBySlides.OrderBy(x => x));
    }

    [Fact]
    public void FindUnusedLayouts_ThreeSlidesSameLayout_ReferencedCountIsThree()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" },
            new TestSlideDefinition { TitleText = "C" });

        var result = Service.FindUnusedLayouts(path);

        var usedLayout = Assert.Single(result.Layouts, l => l.IsUsed);
        Assert.Equal(3, usedLayout.ReferencedBySlides.Count);
    }

    // ──────────────────────────────────────────────────────────
    //  6. Unused layouts exist — extra layouts have IsUsed=false
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_ExtraUnusedLayout_UnusedLayoutCountIsOne()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(1, result.UnusedLayoutCount);
    }

    [Fact]
    public void FindUnusedLayouts_ExtraUnusedLayout_UnusedLayoutIsUsedFalse()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        var unused = result.Layouts.Where(l => !l.IsUsed).ToList();
        Assert.Single(unused);
        Assert.False(unused[0].IsUsed);
    }

    [Fact]
    public void FindUnusedLayouts_ExtraUnusedLayout_UnusedLayoutHasEmptyReferencedBySlides()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        var unused = Assert.Single(result.Layouts, l => !l.IsUsed);
        Assert.Empty(unused.ReferencedBySlides);
    }

    [Fact]
    public void FindUnusedLayouts_ExtraUnusedLayout_TotalLayoutsIsTwo()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(2, result.TotalLayouts);
    }

    // ──────────────────────────────────────────────────────────
    //  7. Estimated savings — sum of unused part sizes
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_AllUsed_EstimatedSavingsIsZero()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(0, result.EstimatedSavingsBytes);
    }

    [Fact]
    public void FindUnusedLayouts_UnusedLayoutExists_EstimatedSavingsGreaterThanZero()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        Assert.True(result.EstimatedSavingsBytes > 0,
            "EstimatedSavingsBytes should be positive when unused layouts exist.");
    }

    [Fact]
    public void FindUnusedLayouts_EstimatedSavings_EqualsUnusedLayoutSizeSum()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        long unusedLayoutBytes = result.Layouts.Where(l => !l.IsUsed).Sum(l => l.SizeBytes);
        long unusedMasterBytes = result.Masters.Where(m => !m.IsUsed).Sum(m => m.SizeBytes);
        Assert.Equal(unusedLayoutBytes + unusedMasterBytes, result.EstimatedSavingsBytes);
    }

    // ──────────────────────────────────────────────────────────
    //  8. Arithmetic invariants
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_TotalLayouts_EqualsUsedPlusUnused()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        int usedCount = result.Layouts.Count(l => l.IsUsed);
        int unusedCount = result.Layouts.Count(l => !l.IsUsed);
        Assert.Equal(result.TotalLayouts, usedCount + unusedCount);
        Assert.Equal(result.UnusedLayoutCount, unusedCount);
    }

    [Fact]
    public void FindUnusedLayouts_TotalMasters_EqualsUsedPlusUnused()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        int usedCount = result.Masters.Count(m => m.IsUsed);
        int unusedCount = result.Masters.Count(m => !m.IsUsed);
        Assert.Equal(result.TotalMasters, usedCount + unusedCount);
        Assert.Equal(result.UnusedMasterCount, unusedCount);
    }

    [Fact]
    public void FindUnusedLayouts_LayoutsCount_EqualsTotalLayouts()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(result.TotalLayouts, result.Layouts.Count);
    }

    [Fact]
    public void FindUnusedLayouts_MastersCount_EqualsTotalMasters()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(result.TotalMasters, result.Masters.Count);
    }

    [Fact]
    public void FindUnusedLayouts_MasterLayoutCount_EqualsLayoutsUnderThatMaster()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        foreach (var master in result.Masters)
        {
            int layoutsUnderMaster = result.Layouts.Count(l => l.MasterName == master.Name);
            Assert.Equal(master.LayoutCount, layoutsUnderMaster);
        }
    }

    [Fact]
    public void FindUnusedLayouts_MasterUsedLayoutCount_MatchesUsedLayoutsUnderThatMaster()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        foreach (var master in result.Masters)
        {
            int usedUnderMaster = result.Layouts.Count(l => l.MasterName == master.Name && l.IsUsed);
            Assert.Equal(master.UsedLayoutCount, usedUnderMaster);
        }
    }

    // ──────────────────────────────────────────────────────────
    //  9. Metadata quality
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_AllMasters_HaveNonEmptyNameAndUri()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.All(result.Masters, m =>
        {
            Assert.False(string.IsNullOrWhiteSpace(m.Name), "Master Name should not be empty.");
            Assert.False(string.IsNullOrWhiteSpace(m.Uri), "Master Uri should not be empty.");
        });
    }

    [Fact]
    public void FindUnusedLayouts_AllLayouts_HaveNonEmptyNameAndUri()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.All(result.Layouts, l =>
        {
            Assert.False(string.IsNullOrWhiteSpace(l.Name), "Layout Name should not be empty.");
            Assert.False(string.IsNullOrWhiteSpace(l.Uri), "Layout Uri should not be empty.");
        });
    }

    [Fact]
    public void FindUnusedLayouts_AllMasters_SizeBytesNonNegative()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.All(result.Masters, m => Assert.True(m.SizeBytes >= 0,
            $"Master '{m.Name}' has negative SizeBytes: {m.SizeBytes}"));
    }

    [Fact]
    public void FindUnusedLayouts_AllLayouts_SizeBytesNonNegative()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.All(result.Layouts, l => Assert.True(l.SizeBytes >= 0,
            $"Layout '{l.Name}' has negative SizeBytes: {l.SizeBytes}"));
    }

    [Fact]
    public void FindUnusedLayouts_FilePath_MatchesInput()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(path, result.FilePath);
    }

    [Fact]
    public void FindUnusedLayouts_Message_IsNotEmpty()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.False(string.IsNullOrWhiteSpace(result.Message));
    }

    [Fact]
    public void FindUnusedLayouts_Warnings_IsNotNull()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.NotNull(result.Warnings);
    }

    // ──────────────────────────────────────────────────────────
    //  10. Master-layout relationship
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_EveryLayout_MasterNameMatchesAMaster()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        var masterNames = result.Masters.Select(m => m.Name).ToHashSet();
        Assert.All(result.Layouts, l =>
            Assert.Contains(l.MasterName, masterNames));
    }

    [Fact]
    public void FindUnusedLayouts_MinimalPptx_LayoutMasterNameMatchesMasterName()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        var master = Assert.Single(result.Masters);
        var layout = Assert.Single(result.Layouts);
        Assert.Equal(master.Name, layout.MasterName);
    }

    // ──────────────────────────────────────────────────────────
    //  11. File not found — throws exception
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_FileNotFound_ThrowsException()
    {
        var bogusPath = Path.Join(Path.GetTempPath(), "nonexistent_" + Guid.NewGuid() + ".pptx");

        Assert.ThrowsAny<Exception>(() => Service.FindUnusedLayouts(bogusPath));
    }

    // ──────────────────────────────────────────────────────────
    //  12. No unused — graceful zero counts
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_NoUnused_UnusedLayoutCountIsZero()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(0, result.UnusedLayoutCount);
    }

    [Fact]
    public void FindUnusedLayouts_NoUnused_UnusedMasterCountIsZero()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(0, result.UnusedMasterCount);
    }

    [Fact]
    public void FindUnusedLayouts_NoUnused_EstimatedSavingsIsZero()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(0, result.EstimatedSavingsBytes);
    }

    [Fact]
    public void FindUnusedLayouts_NoUnused_MessageIndicatesAllUsed()
    {
        var path = CreateMinimalPptx();

        var result = Service.FindUnusedLayouts(path);

        Assert.Contains("All masters and layouts are in use", result.Message);
    }

    // ──────────────────────────────────────────────────────────
    //  Extra: unused master (no slides reference any of its layouts)
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void FindUnusedLayouts_UnusedMaster_MasterIsUsedFalse()
    {
        var path = CreatePptxWithExtraMaster();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(2, result.TotalMasters);
        var unusedMaster = result.Masters.FirstOrDefault(m => !m.IsUsed);
        Assert.NotNull(unusedMaster);
    }

    [Fact]
    public void FindUnusedLayouts_UnusedMaster_UnusedMasterCountIsOne()
    {
        var path = CreatePptxWithExtraMaster();

        var result = Service.FindUnusedLayouts(path);

        Assert.Equal(1, result.UnusedMasterCount);
    }

    [Fact]
    public void FindUnusedLayouts_UnusedMaster_EstimatedSavingsIncludesMasterSize()
    {
        var path = CreatePptxWithExtraMaster();

        var result = Service.FindUnusedLayouts(path);

        var unusedMaster = result.Masters.First(m => !m.IsUsed);
        // Savings includes the master itself and its unused layouts
        Assert.True(result.EstimatedSavingsBytes >= unusedMaster.SizeBytes);
    }

    [Fact]
    public void FindUnusedLayouts_WithUnused_MessageIncludesUnusedCounts()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.FindUnusedLayouts(path);

        Assert.Contains("unused", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ──────────────────────────────────────────────────────────
    //  Helpers — create PPTX fixtures with specific layout structures
    // ──────────────────────────────────────────────────────────

    /// <summary>
    /// Creates a PPTX with 1 master, 2 layouts (Title Slide + Blank), and 1 slide using only Title Slide.
    /// The Blank layout is unused.
    /// </summary>
    private string CreatePptxWithExtraLayout()
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();

        // Layout 1: Title Slide (will be used by the slide)
        var titleLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        titleLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()))
        { Type = SlideLayoutValues.Title };
        titleLayoutPart.SlideLayout.CommonSlideData!.Name = "Title Slide";

        // Layout 2: Blank (unused)
        var blankLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        blankLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()))
        { Type = SlideLayoutValues.Blank };
        blankLayoutPart.SlideLayout.CommonSlideData!.Name = "Blank";

        // Wire layouts back to master
        titleLayoutPart.AddPart(slideMasterPart);
        blankLayoutPart.AddPart(slideMasterPart);

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
                    RelationshipId = slideMasterPart.GetIdOfPart(titleLayoutPart)
                },
                new SlideLayoutId
                {
                    Id = 2050,
                    RelationshipId = slideMasterPart.GetIdOfPart(blankLayoutPart)
                }));

        // One slide using the title layout
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.AddPart(titleLayoutPart);
        slidePart.Slide = new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))));

        var slideIdList = new SlideIdList(
            new SlideId { Id = 256, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

        var slideMasterIdList = new SlideMasterIdList(
            new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

        presentationPart.Presentation = new Presentation(
            slideIdList,
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });

        presentationPart.Presentation.InsertAt(slideMasterIdList, 0);
        presentationPart.Presentation.Save();

        return path;
    }

    /// <summary>
    /// Creates a PPTX with 2 masters: first master has 1 used layout, second master has 1 unused layout.
    /// The slide references only the first master's layout.
    /// </summary>
    private string CreatePptxWithExtraMaster()
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();

        // Master 1 with one layout (used)
        var master1Part = presentationPart.AddNewPart<SlideMasterPart>();
        var layout1Part = master1Part.AddNewPart<SlideLayoutPart>();
        layout1Part.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()))
        { Type = SlideLayoutValues.Title };
        layout1Part.SlideLayout.CommonSlideData!.Name = "Title Slide";
        layout1Part.AddPart(master1Part);

        master1Part.SlideMaster = new SlideMaster(
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
                    RelationshipId = master1Part.GetIdOfPart(layout1Part)
                }));
        master1Part.SlideMaster.CommonSlideData!.Name = "Master 1";

        // Master 2 with one layout (unused)
        var master2Part = presentationPart.AddNewPart<SlideMasterPart>();
        var layout2Part = master2Part.AddNewPart<SlideLayoutPart>();
        layout2Part.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()))
        { Type = SlideLayoutValues.Blank };
        layout2Part.SlideLayout.CommonSlideData!.Name = "Blank Layout";
        layout2Part.AddPart(master2Part);

        master2Part.SlideMaster = new SlideMaster(
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
                    Id = 2051,
                    RelationshipId = master2Part.GetIdOfPart(layout2Part)
                }));
        master2Part.SlideMaster.CommonSlideData!.Name = "Master 2";

        // One slide using layout from Master 1
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.AddPart(layout1Part);
        slidePart.Slide = new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))));

        var slideIdList = new SlideIdList(
            new SlideId { Id = 256, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

        var slideMasterIdList = new SlideMasterIdList(
            new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(master1Part)
            },
            new SlideMasterId
            {
                Id = 2147483649U,
                RelationshipId = presentationPart.GetIdOfPart(master2Part)
            });

        presentationPart.Presentation = new Presentation(
            slideIdList,
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });

        presentationPart.Presentation.InsertAt(slideMasterIdList, 0);
        presentationPart.Presentation.Save();

        return path;
    }
}
