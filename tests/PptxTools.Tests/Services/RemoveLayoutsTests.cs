using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests.Services;

/// <summary>
/// Service-level tests for RemoveUnusedLayouts (Issue #83 — Remove unused slide masters and layouts).
/// Written proactively while Cheritto implements the tool.
/// Validates removal correctness, structural integrity, validation, and round-trip safety.
/// </summary>
[Trait("Category", "Unit")]
public class RemoveLayoutsTests : PptxTestBase
{
    // ──────────────────────────────────────────────────────────
    //  1. Removes unused layout — count decreases
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void RemoveUnusedLayouts_RemovesUnusedLayout()
    {
        var path = CreatePptxWithExtraLayout();

        // Before: 1 master, 2 layouts (1 used, 1 unused)
        var analysisBefore = Service.FindUnusedLayouts(path);
        Assert.Equal(2, analysisBefore.TotalLayouts);
        Assert.Equal(1, analysisBefore.UnusedLayoutCount);

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);
        Assert.True(result.LayoutsRemoved >= 1, "Should remove at least 1 unused layout.");

        // Re-analyze: unused layout should be gone
        var analysisAfter = Service.FindUnusedLayouts(path);
        Assert.Equal(0, analysisAfter.UnusedLayoutCount);
        Assert.True(analysisAfter.TotalLayouts < analysisBefore.TotalLayouts,
            "Total layouts should decrease after removal.");
    }

    // ──────────────────────────────────────────────────────────
    //  2. Preserves used layout — slides still reference it
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void RemoveUnusedLayouts_PreservesUsedLayout()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);

        // The used layout must survive
        var analysisAfter = Service.FindUnusedLayouts(path);
        Assert.True(analysisAfter.TotalLayouts >= 1, "At least 1 used layout must remain.");

        var usedLayout = Assert.Single(analysisAfter.Layouts, l => l.IsUsed);
        Assert.Contains(1, usedLayout.ReferencedBySlides);
    }

    // ──────────────────────────────────────────────────────────
    //  3. Removes unused master — master with zero remaining layouts
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void RemoveUnusedLayouts_RemovesUnusedMaster()
    {
        var path = CreatePptxWithExtraMaster();

        // Before: 2 masters, master 2 is entirely unused
        var analysisBefore = Service.FindUnusedLayouts(path);
        Assert.Equal(2, analysisBefore.TotalMasters);
        Assert.Equal(1, analysisBefore.UnusedMasterCount);

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);
        Assert.True(result.MastersRemoved >= 1, "Should remove the unused master.");

        // Re-analyze: only 1 master should remain
        var analysisAfter = Service.FindUnusedLayouts(path);
        Assert.Equal(1, analysisAfter.TotalMasters);
        Assert.Equal(0, analysisAfter.UnusedMasterCount);
    }

    // ──────────────────────────────────────────────────────────
    //  4. Preserves master with used layouts
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void RemoveUnusedLayouts_PreservesMasterWithUsedLayouts()
    {
        var path = CreatePptxWithExtraMaster();

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);

        // Master 1 (with the used layout) must survive
        var analysisAfter = Service.FindUnusedLayouts(path);
        var usedMaster = Assert.Single(analysisAfter.Masters);
        Assert.True(usedMaster.IsUsed, "The master with used layouts must be preserved.");
        Assert.True(usedMaster.UsedLayoutCount >= 1);
    }

    // ──────────────────────────────────────────────────────────
    //  5. Targeted removal — pass specific URIs, only those removed
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void RemoveUnusedLayouts_TargetedRemoval()
    {
        var path = CreatePptxWithTwoUnusedLayouts();

        // Identify unused layouts first
        var analysis = Service.FindUnusedLayouts(path);
        var unusedLayouts = analysis.Layouts.Where(l => !l.IsUsed).ToList();
        Assert.True(unusedLayouts.Count >= 2, "Fixture should have at least 2 unused layouts.");

        // Target only the first unused layout for removal
        var targetUri = unusedLayouts[0].Uri;
        var result = Service.RemoveUnusedLayouts(path, [targetUri]);

        Assert.True(result.Success);
        Assert.Equal(1, result.LayoutsRemoved);

        // The other unused layout should still be present
        var analysisAfter = Service.FindUnusedLayouts(path);
        Assert.True(analysisAfter.UnusedLayoutCount >= 1,
            "Non-targeted unused layout should still exist.");
    }

    // ──────────────────────────────────────────────────────────
    //  6. Returns bytes saved — positive value
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void RemoveUnusedLayouts_ReturnsSpaceSaved()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);
        Assert.True(result.BytesSaved > 0,
            "BytesSaved should be positive when unused layouts are removed.");
    }

    [Fact]
    public void RemoveUnusedLayouts_BytesSaved_EqualsSumOfRemovedItemSizes()
    {
        var path = CreatePptxWithExtraMaster();

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);
        long sumOfItems = result.RemovedItems.Sum(item => item.SizeBytes);
        Assert.Equal(sumOfItems, result.BytesSaved);
    }

    // ──────────────────────────────────────────────────────────
    //  7. Validates before and after — check validation in output
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void RemoveUnusedLayouts_ValidatesBeforeAndAfter()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);
        Assert.NotNull(result.Validation);
        // Validation errors before should be non-negative
        Assert.True(result.Validation.ErrorsBefore >= 0);
        // After removal, errors should not increase
        Assert.True(result.Validation.ErrorsAfter >= 0);
    }

    [Fact]
    public void RemoveUnusedLayouts_ValidationAfter_DoesNotIntroduceNewErrors()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);
        Assert.True(result.Validation.ErrorsAfter <= result.Validation.ErrorsBefore,
            "Removal should not introduce new validation errors.");
    }

    // ──────────────────────────────────────────────────────────
    //  8. No unused layouts — returns no-op result
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void RemoveUnusedLayouts_NoUnusedLayouts_ReturnsNoOp()
    {
        var path = CreateMinimalPptx();

        // Pre-check: all layouts used
        var analysis = Service.FindUnusedLayouts(path);
        Assert.Equal(0, analysis.UnusedLayoutCount);

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);
        Assert.Equal(0, result.LayoutsRemoved);
        Assert.Equal(0, result.MastersRemoved);
        Assert.Equal(0, result.BytesSaved);
        Assert.Empty(result.RemovedItems);
    }

    // ──────────────────────────────────────────────────────────
    //  9. Invalid file path — returns error
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void RemoveUnusedLayouts_InvalidFilePath_ReturnsError()
    {
        var bogusPath = Path.Join(Path.GetTempPath(), "nonexistent_" + Guid.NewGuid() + ".pptx");

        Assert.ThrowsAny<Exception>(() => Service.RemoveUnusedLayouts(bogusPath));
    }

    // ──────────────────────────────────────────────────────────
    //  10. Round-trip — file opens in OpenXml after removal
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void RemoveUnusedLayouts_FileOpensInOpenXml_AfterRemoval()
    {
        var path = CreatePptxWithExtraMaster();

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);

        // Re-open and verify structural integrity
        using var doc = PresentationDocument.Open(path, false);
        var presentationPart = doc.PresentationPart;
        Assert.NotNull(presentationPart);

        // Presentation should still have slides
        var slideIdList = presentationPart.Presentation.SlideIdList;
        Assert.NotNull(slideIdList);
        Assert.NotEmpty(slideIdList.Elements<SlideId>());

        // At least one master and one layout must remain
        Assert.NotEmpty(presentationPart.SlideMasterParts);
        var remainingMaster = presentationPart.SlideMasterParts.First();
        Assert.NotEmpty(remainingMaster.SlideLayoutParts);

        // SlideMasterIdList should be consistent with actual master parts
        var masterIdList = presentationPart.Presentation.SlideMasterIdList;
        Assert.NotNull(masterIdList);
        int masterIdCount = masterIdList.Elements<SlideMasterId>().Count();
        int actualMasterCount = presentationPart.SlideMasterParts.Count();
        Assert.Equal(actualMasterCount, masterIdCount);
    }

    [Fact]
    public void RemoveUnusedLayouts_SlidesStillAccessible_AfterRemoval()
    {
        var path = CreatePptxWithExtraMaster();

        var result = Service.RemoveUnusedLayouts(path);
        Assert.True(result.Success);

        // Verify each slide can still reach its layout and master
        using var doc = PresentationDocument.Open(path, false);
        var presentationPart = doc.PresentationPart!;

        foreach (var slideId in presentationPart.Presentation.SlideIdList!.Elements<SlideId>())
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!);
            Assert.NotNull(slidePart.Slide);
            Assert.NotNull(slidePart.SlideLayoutPart);
            Assert.NotNull(slidePart.SlideLayoutPart.SlideMasterPart);
        }
    }

    // ──────────────────────────────────────────────────────────
    //  Additional: result metadata quality
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void RemoveUnusedLayouts_FilePath_MatchesInput()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.RemoveUnusedLayouts(path);

        Assert.Equal(path, result.FilePath);
    }

    [Fact]
    public void RemoveUnusedLayouts_Message_IsNotEmpty()
    {
        var path = CreatePptxWithExtraLayout();

        var result = Service.RemoveUnusedLayouts(path);

        Assert.False(string.IsNullOrWhiteSpace(result.Message));
    }

    [Fact]
    public void RemoveUnusedLayouts_RemovedItems_HaveValidMetadata()
    {
        var path = CreatePptxWithExtraMaster();

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);
        Assert.NotEmpty(result.RemovedItems);

        Assert.All(result.RemovedItems, item =>
        {
            Assert.False(string.IsNullOrWhiteSpace(item.Name), "Removed item Name should not be empty.");
            Assert.False(string.IsNullOrWhiteSpace(item.Uri), "Removed item Uri should not be empty.");
            Assert.Contains(item.Type, new[] { "layout", "master" });
            Assert.True(item.SizeBytes >= 0, $"Removed item '{item.Name}' should have non-negative size.");
        });
    }

    [Fact]
    public void RemoveUnusedLayouts_RemovedItems_ContainsBothLayoutAndMaster()
    {
        var path = CreatePptxWithExtraMaster();

        var result = Service.RemoveUnusedLayouts(path);

        Assert.True(result.Success);
        Assert.Contains(result.RemovedItems, item => item.Type == "layout");
        Assert.Contains(result.RemovedItems, item => item.Type == "master");
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

    /// <summary>
    /// Creates a PPTX with 1 master, 3 layouts (Title + Blank + Section Header),
    /// and 1 slide using only Title. Both Blank and Section Header are unused.
    /// Used for targeted removal tests.
    /// </summary>
    private string CreatePptxWithTwoUnusedLayouts()
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();

        // Layout 1: Title Slide (used)
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
        titleLayoutPart.AddPart(slideMasterPart);

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
        blankLayoutPart.AddPart(slideMasterPart);

        // Layout 3: Section Header (unused)
        var sectionLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        sectionLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()))
        { Type = SlideLayoutValues.SectionHeader };
        sectionLayoutPart.SlideLayout.CommonSlideData!.Name = "Section Header";
        sectionLayoutPart.AddPart(slideMasterPart);

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
                },
                new SlideLayoutId
                {
                    Id = 2051,
                    RelationshipId = slideMasterPart.GetIdOfPart(sectionLayoutPart)
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
}
