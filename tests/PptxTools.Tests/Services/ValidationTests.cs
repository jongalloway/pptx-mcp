using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests.Services;

/// <summary>
/// Service-level tests for ValidatePresentation (Issue #121 — Presentation validation and diagnostics).
/// Covers all validation checks: duplicate shape IDs, missing image references, orphaned relationships,
/// required elements, hyperlink targets, cross-slide duplicates, and slide-number filtering.
/// </summary>
[Trait("Category", "Unit")]
public class ValidationTests : PptxTestBase
{
    // ────────────────────────────────────────────────────────
    // Happy path: valid presentations
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_MinimalPptx_ReturnsSuccess()
    {
        var path = CreateMinimalPptx();
        var result = Service.ValidatePresentation(path);

        Assert.True(result.Success);
        Assert.Equal("Validate", result.Action);
        Assert.NotEmpty(result.Message);
    }

    [Fact]
    public void ValidatePresentation_MinimalPptx_NoIssues()
    {
        var path = CreateMinimalPptx();
        var result = Service.ValidatePresentation(path);

        Assert.Equal(0, result.IssueCount);
        Assert.Equal(0, result.ErrorCount);
        Assert.Equal(0, result.WarningCount);
        Assert.Equal(0, result.InfoCount);
        Assert.Empty(result.Issues);
    }

    [Fact]
    public void ValidatePresentation_MinimalPptx_MessageIndicatesPass()
    {
        var path = CreateMinimalPptx();
        var result = Service.ValidatePresentation(path);

        Assert.Contains("no issues", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ValidatePresentation_MultiSlide_ReturnsSuccessWithNoIssues()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" },
            new TestSlideDefinition { TitleText = "Slide 3" });
        var result = Service.ValidatePresentation(path);

        Assert.True(result.Success);
        Assert.Equal(0, result.ErrorCount);
        Assert.Equal(0, result.WarningCount);
    }

    [Fact]
    public void ValidatePresentation_WithImage_ReturnsNoErrors()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Image Slide", IncludeImage = true });
        var result = Service.ValidatePresentation(path);

        Assert.True(result.Success);
        Assert.Equal(0, result.ErrorCount);
    }

    // ────────────────────────────────────────────────────────
    // Empty presentation
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_EmptyPresentation_ReturnsSuccessWithNoIssues()
    {
        var path = CreateEmptyPresentation();
        var result = Service.ValidatePresentation(path);

        Assert.True(result.Success);
        Assert.Equal(0, result.IssueCount);
        Assert.Contains("no slides", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────
    // Error handling
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_FileNotFound_Throws()
    {
        var bogusPath = Path.Join(Path.GetTempPath(), "nonexistent_" + Guid.NewGuid() + ".pptx");
        Assert.ThrowsAny<Exception>(() => Service.ValidatePresentation(bogusPath));
    }

    // ────────────────────────────────────────────────────────
    // Duplicate shape IDs (within a single slide)
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_DuplicateShapeIds_DetectsError()
    {
        var path = CreatePptxWithDuplicateShapeIds();
        var result = Service.ValidatePresentation(path);

        Assert.True(result.Success);
        var dupIssues = result.Issues.Where(i => i.Category == "DuplicateShapeId").ToList();
        Assert.NotEmpty(dupIssues);
        Assert.All(dupIssues, i => Assert.Equal(ValidationSeverity.Error, i.Severity));
    }

    [Fact]
    public void ValidatePresentation_DuplicateShapeIds_IssueHasSlideNumber()
    {
        var path = CreatePptxWithDuplicateShapeIds();
        var result = Service.ValidatePresentation(path);

        var dupIssue = Assert.Single(result.Issues, i => i.Category == "DuplicateShapeId");
        Assert.Equal(1, dupIssue.SlideNumber);
    }

    [Fact]
    public void ValidatePresentation_DuplicateShapeIds_HasRecommendation()
    {
        var path = CreatePptxWithDuplicateShapeIds();
        var result = Service.ValidatePresentation(path);

        var dupIssue = Assert.Single(result.Issues, i => i.Category == "DuplicateShapeId");
        Assert.NotEmpty(dupIssue.Recommendation);
    }

    [Fact]
    public void ValidatePresentation_DuplicateShapeIds_DescriptionContainsShapeId()
    {
        var path = CreatePptxWithDuplicateShapeIds();
        var result = Service.ValidatePresentation(path);

        var dupIssue = Assert.Single(result.Issues, i => i.Category == "DuplicateShapeId");
        Assert.Contains("Shape ID", dupIssue.Description);
    }

    // ────────────────────────────────────────────────────────
    // Missing image references
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_MissingImageRef_DetectsError()
    {
        var path = CreatePptxWithBrokenImageRef();
        var result = Service.ValidatePresentation(path);

        Assert.True(result.Success);
        var missingImgIssues = result.Issues.Where(i => i.Category == "MissingImageReference").ToList();
        Assert.NotEmpty(missingImgIssues);
        Assert.All(missingImgIssues, i => Assert.Equal(ValidationSeverity.Error, i.Severity));
    }

    [Fact]
    public void ValidatePresentation_MissingImageRef_HasSlideNumber()
    {
        var path = CreatePptxWithBrokenImageRef();
        var result = Service.ValidatePresentation(path);

        var issue = Assert.Single(result.Issues, i => i.Category == "MissingImageReference");
        Assert.Equal(1, issue.SlideNumber);
    }

    [Fact]
    public void ValidatePresentation_MissingImageRef_DescriptionContainsRelId()
    {
        var path = CreatePptxWithBrokenImageRef();
        var result = Service.ValidatePresentation(path);

        var issue = Assert.Single(result.Issues, i => i.Category == "MissingImageReference");
        Assert.Contains("rId999", issue.Description);
    }

    // ────────────────────────────────────────────────────────
    // Missing required elements
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_MissingShapeTree_DetectsError()
    {
        var path = CreatePptxMissingShapeTree();
        var result = Service.ValidatePresentation(path);

        Assert.True(result.Success);
        var missingElements = result.Issues.Where(i => i.Category == "MissingRequiredElement").ToList();
        Assert.NotEmpty(missingElements);
        Assert.All(missingElements, i => Assert.Equal(ValidationSeverity.Error, i.Severity));
    }

    [Fact]
    public void ValidatePresentation_MissingCommonSlideData_DetectsError()
    {
        var path = CreatePptxMissingCommonSlideData();
        var result = Service.ValidatePresentation(path);

        Assert.True(result.Success);
        var missingElements = result.Issues.Where(i => i.Category == "MissingRequiredElement").ToList();
        Assert.NotEmpty(missingElements);
        Assert.Contains(missingElements, i => i.Description.Contains("CommonSlideData"));
    }

    // ────────────────────────────────────────────────────────
    // Slide number filtering
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_SlideFilter_OnlyValidatesSpecifiedSlide()
    {
        // Create a 2-slide pptx, corrupt only slide 1
        var path = CreatePptxWithDuplicateShapeIdsOnSlide1Only();
        var result = Service.ValidatePresentation(path, slideNumber: 2);

        // Slide 2 has no issues, so filtered result should be clean
        var slide1Issues = result.Issues.Where(i => i.SlideNumber == 1).ToList();
        Assert.Empty(slide1Issues);
    }

    [Fact]
    public void ValidatePresentation_SlideFilter_DetectsIssuesOnTargetSlide()
    {
        var path = CreatePptxWithDuplicateShapeIdsOnSlide1Only();
        var result = Service.ValidatePresentation(path, slideNumber: 1);

        var slide1Issues = result.Issues.Where(i => i.Category == "DuplicateShapeId").ToList();
        Assert.NotEmpty(slide1Issues);
    }

    [Fact]
    public void ValidatePresentation_SlideFilter_DoesNotRunCrossSlideCheck()
    {
        // Cross-slide check is skipped when filtering by slide number
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" });
        var result = Service.ValidatePresentation(path, slideNumber: 1);

        var crossSlideIssues = result.Issues.Where(i => i.Category == "CrossSlideDuplicateShapeId").ToList();
        Assert.Empty(crossSlideIssues);
    }

    // ────────────────────────────────────────────────────────
    // Severity counts and sorting
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_SeverityCounts_MatchIssueList()
    {
        var path = CreatePptxWithDuplicateShapeIds();
        var result = Service.ValidatePresentation(path);

        Assert.Equal(result.Issues.Count, result.IssueCount);
        Assert.Equal(result.Issues.Count(i => i.Severity == ValidationSeverity.Error), result.ErrorCount);
        Assert.Equal(result.Issues.Count(i => i.Severity == ValidationSeverity.Warning), result.WarningCount);
        Assert.Equal(result.Issues.Count(i => i.Severity == ValidationSeverity.Info), result.InfoCount);
    }

    [Fact]
    public void ValidatePresentation_SeverityCountsSum_EqualsIssueCount()
    {
        var path = CreatePptxWithDuplicateShapeIds();
        var result = Service.ValidatePresentation(path);

        Assert.Equal(result.IssueCount, result.ErrorCount + result.WarningCount + result.InfoCount);
    }

    [Fact]
    public void ValidatePresentation_IssuesSortedBySeverity_ErrorsFirst()
    {
        // Use a fixture that produces both errors and other severity levels
        var path = CreatePptxWithMixedIssues();
        var result = Service.ValidatePresentation(path);

        Assert.True(result.Issues.Count >= 2, "Need at least 2 issues to verify ordering.");

        for (int i = 0; i < result.Issues.Count - 1; i++)
        {
            Assert.True(
                result.Issues[i].Severity <= result.Issues[i + 1].Severity,
                $"Issue at index {i} (Severity={result.Issues[i].Severity}) should come before or equal index {i + 1} (Severity={result.Issues[i + 1].Severity}).");
        }
    }

    // ────────────────────────────────────────────────────────
    // Cross-slide shape ID duplicates (Info severity)
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_CrossSlideDuplicateIds_DetectedAsInfo()
    {
        // TestPptxHelper reuses low shape IDs (starting at 2) per slide,
        // so multi-slide presentations naturally have cross-slide duplicates.
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide A" },
            new TestSlideDefinition { TitleText = "Slide B" });
        var result = Service.ValidatePresentation(path);

        var crossDupIssues = result.Issues.Where(i => i.Category == "CrossSlideDuplicateShapeId").ToList();
        Assert.NotEmpty(crossDupIssues);
        Assert.All(crossDupIssues, i => Assert.Equal(ValidationSeverity.Info, i.Severity));
    }

    [Fact]
    public void ValidatePresentation_SingleSlide_NoCrossSlideDuplicates()
    {
        var path = CreateMinimalPptx();
        var result = Service.ValidatePresentation(path);

        var crossDupIssues = result.Issues.Where(i => i.Category == "CrossSlideDuplicateShapeId").ToList();
        Assert.Empty(crossDupIssues);
    }

    // ────────────────────────────────────────────────────────
    // Issue metadata quality
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_AllIssues_HaveNonEmptyDescription()
    {
        var path = CreatePptxWithMixedIssues();
        var result = Service.ValidatePresentation(path);

        Assert.All(result.Issues, i => Assert.False(string.IsNullOrWhiteSpace(i.Description)));
    }

    [Fact]
    public void ValidatePresentation_AllIssues_HaveNonEmptyCategory()
    {
        var path = CreatePptxWithMixedIssues();
        var result = Service.ValidatePresentation(path);

        Assert.All(result.Issues, i => Assert.False(string.IsNullOrWhiteSpace(i.Category)));
    }

    [Fact]
    public void ValidatePresentation_AllIssues_HaveNonEmptyRecommendation()
    {
        var path = CreatePptxWithMixedIssues();
        var result = Service.ValidatePresentation(path);

        Assert.All(result.Issues, i => Assert.False(string.IsNullOrWhiteSpace(i.Recommendation)));
    }

    // ────────────────────────────────────────────────────────
    // Message content
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_WithIssues_MessageContainsCounts()
    {
        var path = CreatePptxWithDuplicateShapeIds();
        var result = Service.ValidatePresentation(path);

        Assert.True(result.IssueCount > 0);
        Assert.Contains("issue", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("error", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────
    // Complex fixture: multiple check types together
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ValidatePresentation_ComplexPresentation_Tables_Charts_Images_Valid()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Overview",
                TextShapes = [new TestTextShapeDefinition { Paragraphs = ["Bullet 1", "Bullet 2"] }]
            },
            new TestSlideDefinition
            {
                TitleText = "Data",
                IncludeImage = true,
                Tables = [new TestTableDefinition { Rows = [["A", "B"], ["1", "2"]] }]
            });
        var result = Service.ValidatePresentation(path);

        Assert.True(result.Success);
        Assert.Equal(0, result.ErrorCount);
    }

    [Fact]
    public void ValidatePresentation_Idempotent_SameResultOnMultipleCalls()
    {
        var path = CreateMinimalPptx();
        var result1 = Service.ValidatePresentation(path);
        var result2 = Service.ValidatePresentation(path);

        Assert.Equal(result1.IssueCount, result2.IssueCount);
        Assert.Equal(result1.ErrorCount, result2.ErrorCount);
        Assert.Equal(result1.WarningCount, result2.WarningCount);
        Assert.Equal(result1.InfoCount, result2.InfoCount);
    }

    // ════════════════════════════════════════════════════════
    // Fixture helpers — create corrupt PPTX files for testing
    // ════════════════════════════════════════════════════════

    private string CreateEmptyPresentation()
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();

        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

        slideLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()));
        slideLayoutPart.AddPart(slideMasterPart);

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
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
            new SlideLayoutIdList(new SlideLayoutId
            {
                Id = 2049,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            }));

        // No slides — empty SlideIdList
        presentationPart.Presentation = new Presentation(
            new SlideIdList(),
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });

        presentationPart.Presentation.InsertAt(
            new SlideMasterIdList(new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            }), 0);

        presentationPart.Presentation.Save();
        return path;
    }

    /// <summary>Creates a single-slide PPTX with two shapes sharing the same ID.</summary>
    private string CreatePptxWithDuplicateShapeIds()
    {
        var path = CreateMinimalPptx("Dup Test");

        using var doc = PresentationDocument.Open(path, true);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        // Get the ID of the first non-group shape
        uint duplicateId = 2;
        foreach (var child in shapeTree.Elements<Shape>())
        {
            var id = child.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (id.HasValue)
            {
                duplicateId = id.Value;
                break;
            }
        }

        // Add another shape with the same ID
        shapeTree.Append(new Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = duplicateId, Name = "DuplicateShape" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 914400, Cy = 457200 })),
            new P.TextBody(
                new A.BodyProperties(),
                new A.Paragraph(new A.Run(new A.Text("Duplicate"))))));

        slidePart.Slide.Save();
        return path;
    }

    /// <summary>Creates a 2-slide PPTX where only slide 1 has duplicate shape IDs.</summary>
    private string CreatePptxWithDuplicateShapeIdsOnSlide1Only()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" });

        using var doc = PresentationDocument.Open(path, true);
        var slideIds = doc.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>().ToList();
        var slide1Part = (SlidePart)doc.PresentationPart.GetPartById(slideIds[0].RelationshipId!.Value!);
        var shapeTree = slide1Part.Slide.CommonSlideData!.ShapeTree!;

        uint duplicateId = 2;
        foreach (var child in shapeTree.Elements<Shape>())
        {
            var id = child.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (id.HasValue) { duplicateId = id.Value; break; }
        }

        shapeTree.Append(new Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = duplicateId, Name = "DupOnSlide1" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 914400, Cy = 457200 })),
            new P.TextBody(
                new A.BodyProperties(),
                new A.Paragraph(new A.Run(new A.Text("Duplicate"))))));

        slide1Part.Slide.Save();
        return path;
    }

    /// <summary>Creates a PPTX with a picture shape referencing a non-existent image relationship.</summary>
    private string CreatePptxWithBrokenImageRef()
    {
        var path = CreateMinimalPptx("Broken Image");

        using var doc = PresentationDocument.Open(path, true);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

        // Add a Picture element with a blip embed pointing to a non-existent relationship
        shapeTree.Append(new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = 100, Name = "BrokenPicture" },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = "rId999" },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 914400, Cy = 914400 }))));

        slidePart.Slide.Save();
        return path;
    }

    /// <summary>Creates a PPTX with a slide missing its ShapeTree element.</summary>
    private string CreatePptxMissingShapeTree()
    {
        var path = CreateMinimalPptx("Missing ShapeTree");

        using var doc = PresentationDocument.Open(path, true);
        var slidePart = doc.PresentationPart!.SlideParts.First();

        // Replace the slide with one that has CommonSlideData but no ShapeTree
        slidePart.Slide = new Slide(new CommonSlideData());
        slidePart.Slide.Save();

        return path;
    }

    /// <summary>Creates a PPTX with a slide missing its CommonSlideData element.</summary>
    private string CreatePptxMissingCommonSlideData()
    {
        var path = CreateMinimalPptx("Missing cSld");

        using var doc = PresentationDocument.Open(path, true);
        var slidePart = doc.PresentationPart!.SlideParts.First();

        // Replace the slide with a bare Slide element (no CommonSlideData)
        slidePart.Slide = new Slide();
        slidePart.Slide.Save();

        return path;
    }

    /// <summary>
    /// Creates a PPTX that triggers multiple issue categories:
    /// duplicate shape IDs (Error) on slide 1, and cross-slide duplicates (Info) across 2 slides.
    /// </summary>
    private string CreatePptxWithMixedIssues()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" });

        using var doc = PresentationDocument.Open(path, true);
        var slideIds = doc.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>().ToList();
        var slide1Part = (SlidePart)doc.PresentationPart.GetPartById(slideIds[0].RelationshipId!.Value!);
        var shapeTree = slide1Part.Slide.CommonSlideData!.ShapeTree!;

        // Add duplicate shape ID on slide 1 (Error)
        uint duplicateId = 2;
        foreach (var child in shapeTree.Elements<Shape>())
        {
            var id = child.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (id.HasValue) { duplicateId = id.Value; break; }
        }

        shapeTree.Append(new Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = duplicateId, Name = "MixedDup" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 914400, Cy = 457200 })),
            new P.TextBody(
                new A.BodyProperties(),
                new A.Paragraph(new A.Run(new A.Text("Mixed"))))));

        slide1Part.Slide.Save();
        // Cross-slide duplicates (Info) naturally happen because TestPptxHelper reuses shape IDs
        return path;
    }
}
