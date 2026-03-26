using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests.Services;

/// <summary>
/// Service-level tests for ComparePresentations (Issue #133 — Compare two presentations).
/// Covers Full, SlidesOnly, TextOnly, MetadataOnly actions, identical/different presentations,
/// slide count mismatches, text changes, metadata changes, and error handling.
/// </summary>
[Trait("Category", "Unit")]
public class CompareTests : PptxTestBase
{
    // ────────────────────────────────────────────────────────
    // Fixture helpers
    // ────────────────────────────────────────────────────────

    /// <summary>Creates a presentation with specific metadata set on the package properties.</summary>
    private string CreatePptxWithMetadata(
        string? title = null,
        string? creator = null,
        string? subject = null,
        string? keywords = null,
        string? description = null,
        string? category = null,
        params TestSlideDefinition[] slides)
    {
        var defs = slides.Length > 0 ? slides : [new TestSlideDefinition { TitleText = "Slide 1" }];
        var path = CreatePptxWithSlides(defs);

        using var doc = PresentationDocument.Open(path, true);
        if (title is not null) doc.PackageProperties.Title = title;
        if (creator is not null) doc.PackageProperties.Creator = creator;
        if (subject is not null) doc.PackageProperties.Subject = subject;
        if (keywords is not null) doc.PackageProperties.Keywords = keywords;
        if (description is not null) doc.PackageProperties.Description = description;
        if (category is not null) doc.PackageProperties.Category = category;

        return path;
    }

    /// <summary>Creates two identical presentations with the same slide structure and metadata.</summary>
    private (string Source, string Target) CreateIdenticalPair(
        string? title = "Test Deck",
        string? creator = "Author")
    {
        var slides = new[]
        {
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" }
        };

        var source = CreatePptxWithMetadata(title: title, creator: creator, slides: slides);
        var target = CreatePptxWithMetadata(title: title, creator: creator, slides: slides);
        return (source, target);
    }

    // ────────────────────────────────────────────────────────
    // Full comparison: identical presentations
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ComparePresentations_Identical_ReturnsSuccess()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.True(result.Success);
    }

    [Fact]
    public void ComparePresentations_Identical_AreIdenticalTrue()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.True(result.AreIdentical);
        Assert.Equal(0, result.DifferenceCount);
    }

    [Fact]
    public void ComparePresentations_Identical_NoSlideDifferences()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.SlideDifferences);
        Assert.Empty(result.SlideDifferences);
    }

    [Fact]
    public void ComparePresentations_Identical_NoTextDifferences()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.TextDifferences);
        Assert.Empty(result.TextDifferences);
    }

    [Fact]
    public void ComparePresentations_Identical_NoMetadataDifferences()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.MetadataDifferences);
        Assert.Empty(result.MetadataDifferences);
    }

    [Fact]
    public void ComparePresentations_Identical_ActionIsFull()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.Equal("Full", result.Action);
    }

    // ────────────────────────────────────────────────────────
    // Slide count differences
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ComparePresentations_DifferentSlideCount_AreIdenticalFalse()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" },
            new TestSlideDefinition { TitleText = "Slide 3" });

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.False(result.AreIdentical);
        Assert.True(result.DifferenceCount > 0);
    }

    [Fact]
    public void ComparePresentations_TargetHasMoreSlides_ReportsSlideDifference()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" });

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.SlideDifferences);
        Assert.NotEmpty(result.SlideDifferences);
    }

    [Fact]
    public void ComparePresentations_SourceHasMoreSlides_ReportsSlideDifference()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" },
            new TestSlideDefinition { TitleText = "Slide 3" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" });

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.SlideDifferences);
        Assert.NotEmpty(result.SlideDifferences);
    }

    [Fact]
    public void ComparePresentations_AddedSlides_MarkedAsAdded()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "New Slide" });

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.SlideDifferences);
        Assert.Contains(result.SlideDifferences, d =>
            d.DifferenceType.Equals("Added", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ComparePresentations_RemovedSlides_MarkedAsRemoved()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" });

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.SlideDifferences);
        Assert.Contains(result.SlideDifferences, d =>
            d.DifferenceType.Equals("Removed", StringComparison.OrdinalIgnoreCase));
    }

    // ────────────────────────────────────────────────────────
    // Text content differences
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ComparePresentations_TextChanged_AreIdenticalFalse()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Original Title" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Modified Title" });

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.False(result.AreIdentical);
    }

    [Fact]
    public void ComparePresentations_TextChanged_ReportsTextDifference()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Original Title" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Modified Title" });

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.TextDifferences);
        Assert.NotEmpty(result.TextDifferences);
    }

    [Fact]
    public void ComparePresentations_TextChanged_IncludesSourceAndTargetText()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Alpha" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Beta" });

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.TextDifferences);
        var diff = Assert.Single(result.TextDifferences.Where(d => d.SlideNumber == 1));
        Assert.Contains("Alpha", diff.SourceText ?? "");
        Assert.Contains("Beta", diff.TargetText ?? "");
    }

    [Fact]
    public void ComparePresentations_MultiSlideTextChanges_ReportsPerSlide()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Same" },
            new TestSlideDefinition { TitleText = "Original" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Same" },
            new TestSlideDefinition { TitleText = "Changed" });

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.TextDifferences);
        // Only the second slide should have text differences
        Assert.All(result.TextDifferences, d => Assert.Equal(2, d.SlideNumber));
    }

    [Fact]
    public void ComparePresentations_BodyTextChanged_ReportsTextDifference()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Same Title",
                TextShapes = [new TestTextShapeDefinition { Name = "Body", Paragraphs = ["Original body text"] }]
            });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Same Title",
                TextShapes = [new TestTextShapeDefinition { Name = "Body", Paragraphs = ["Modified body text"] }]
            });

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.TextDifferences);
        Assert.NotEmpty(result.TextDifferences);
    }

    // ────────────────────────────────────────────────────────
    // Metadata differences
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ComparePresentations_DifferentTitle_AreIdenticalFalse()
    {
        var source = CreatePptxWithMetadata(title: "Deck A");
        var target = CreatePptxWithMetadata(title: "Deck B");

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.False(result.AreIdentical);
    }

    [Fact]
    public void ComparePresentations_DifferentTitle_ReportsMetadataDifference()
    {
        var source = CreatePptxWithMetadata(title: "Deck A");
        var target = CreatePptxWithMetadata(title: "Deck B");

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.MetadataDifferences);
        Assert.Contains(result.MetadataDifferences, d =>
            d.Property.Equals("Title", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ComparePresentations_DifferentCreator_ReportsMetadataDifference()
    {
        var source = CreatePptxWithMetadata(creator: "Alice");
        var target = CreatePptxWithMetadata(creator: "Bob");

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.MetadataDifferences);
        Assert.Contains(result.MetadataDifferences, d =>
            d.Property.Equals("Creator", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ComparePresentations_MetadataDifference_IncludesSourceAndTargetValues()
    {
        var source = CreatePptxWithMetadata(title: "Source Title");
        var target = CreatePptxWithMetadata(title: "Target Title");

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.MetadataDifferences);
        var titleDiff = result.MetadataDifferences.FirstOrDefault(d =>
            d.Property.Equals("Title", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(titleDiff);
        Assert.Equal("Source Title", titleDiff.SourceValue);
        Assert.Equal("Target Title", titleDiff.TargetValue);
    }

    [Fact]
    public void ComparePresentations_MultipleMetadataChanges_ReportsAll()
    {
        var source = CreatePptxWithMetadata(title: "A", creator: "Alice", subject: "X");
        var target = CreatePptxWithMetadata(title: "B", creator: "Bob", subject: "Y");

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotNull(result.MetadataDifferences);
        Assert.True(result.MetadataDifferences.Count >= 3);
    }

    // ────────────────────────────────────────────────────────
    // Action-specific: SlidesOnly
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ComparePresentations_SlidesOnly_ReturnsActionName()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.SlidesOnly);

        Assert.Equal("SlidesOnly", result.Action);
    }

    [Fact]
    public void ComparePresentations_SlidesOnly_IdenticalSlides_NoDifferences()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.SlidesOnly);

        Assert.True(result.AreIdentical);
        Assert.NotNull(result.SlideDifferences);
        Assert.Empty(result.SlideDifferences);
    }

    [Fact]
    public void ComparePresentations_SlidesOnly_DifferentCounts_ReportsDifference()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Only slide" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" });

        var result = Service.ComparePresentations(source, target, CompareAction.SlidesOnly);

        Assert.False(result.AreIdentical);
        Assert.NotNull(result.SlideDifferences);
        Assert.NotEmpty(result.SlideDifferences);
    }

    // ────────────────────────────────────────────────────────
    // Action-specific: TextOnly
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ComparePresentations_TextOnly_ReturnsActionName()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.TextOnly);

        Assert.Equal("TextOnly", result.Action);
    }

    [Fact]
    public void ComparePresentations_TextOnly_IdenticalText_NoDifferences()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.TextOnly);

        Assert.True(result.AreIdentical);
        Assert.NotNull(result.TextDifferences);
        Assert.Empty(result.TextDifferences);
    }

    [Fact]
    public void ComparePresentations_TextOnly_TextChanged_ReportsDifference()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Hello" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "World" });

        var result = Service.ComparePresentations(source, target, CompareAction.TextOnly);

        Assert.False(result.AreIdentical);
        Assert.NotNull(result.TextDifferences);
        Assert.NotEmpty(result.TextDifferences);
    }

    // ────────────────────────────────────────────────────────
    // Action-specific: MetadataOnly
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ComparePresentations_MetadataOnly_ReturnsActionName()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.MetadataOnly);

        Assert.Equal("MetadataOnly", result.Action);
    }

    [Fact]
    public void ComparePresentations_MetadataOnly_IdenticalMetadata_NoDifferences()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.MetadataOnly);

        Assert.True(result.AreIdentical);
        Assert.NotNull(result.MetadataDifferences);
        Assert.Empty(result.MetadataDifferences);
    }

    [Fact]
    public void ComparePresentations_MetadataOnly_DifferentTitle_ReportsDifference()
    {
        var source = CreatePptxWithMetadata(title: "V1");
        var target = CreatePptxWithMetadata(title: "V2");

        var result = Service.ComparePresentations(source, target, CompareAction.MetadataOnly);

        Assert.False(result.AreIdentical);
        Assert.NotNull(result.MetadataDifferences);
        Assert.NotEmpty(result.MetadataDifferences);
    }

    // ────────────────────────────────────────────────────────
    // Error handling: file not found
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ComparePresentations_SourceFileNotFound_Throws()
    {
        var target = CreateMinimalPptx();
        var fakePath = Path.Combine(Path.GetTempPath(), "nonexistent-source.pptx");

        Assert.ThrowsAny<Exception>(() =>
            Service.ComparePresentations(fakePath, target, CompareAction.Full));
    }

    [Fact]
    public void ComparePresentations_TargetFileNotFound_Throws()
    {
        var source = CreateMinimalPptx();
        var fakePath = Path.Combine(Path.GetTempPath(), "nonexistent-target.pptx");

        Assert.ThrowsAny<Exception>(() =>
            Service.ComparePresentations(source, fakePath, CompareAction.Full));
    }

    [Fact]
    public void ComparePresentations_BothFilesNotFound_Throws()
    {
        var fakeSource = Path.Combine(Path.GetTempPath(), "missing-source.pptx");
        var fakeTarget = Path.Combine(Path.GetTempPath(), "missing-target.pptx");

        Assert.ThrowsAny<Exception>(() =>
            Service.ComparePresentations(fakeSource, fakeTarget, CompareAction.Full));
    }

    // ────────────────────────────────────────────────────────
    // Edge cases
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ComparePresentations_SameFile_AreIdentical()
    {
        var path = CreateMinimalPptx();

        var result = Service.ComparePresentations(path, path, CompareAction.Full);

        Assert.True(result.Success);
        Assert.True(result.AreIdentical);
        Assert.Equal(0, result.DifferenceCount);
    }

    [Fact]
    public void ComparePresentations_EmptyPresentations_AreIdentical()
    {
        // Minimal presentations with the same default structure
        var source = CreateMinimalPptx();
        var target = CreateMinimalPptx();

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.True(result.Success);
        Assert.True(result.AreIdentical);
    }

    [Fact]
    public void ComparePresentations_DifferenceCount_MatchesTotalDifferences()
    {
        var source = CreatePptxWithMetadata(
            title: "A", creator: "Alice",
            slides: [new TestSlideDefinition { TitleText = "Hello" }]);
        var target = CreatePptxWithMetadata(
            title: "B", creator: "Bob",
            slides:
            [
                new TestSlideDefinition { TitleText = "World" },
                new TestSlideDefinition { TitleText = "Extra" }
            ]);

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        var expectedTotal =
            (result.SlideDifferences?.Count ?? 0) +
            (result.TextDifferences?.Count ?? 0) +
            (result.MetadataDifferences?.Count ?? 0);

        Assert.Equal(expectedTotal, result.DifferenceCount);
    }

    [Theory]
    [InlineData(CompareAction.Full)]
    [InlineData(CompareAction.SlidesOnly)]
    [InlineData(CompareAction.TextOnly)]
    [InlineData(CompareAction.MetadataOnly)]
    public void ComparePresentations_AllActions_ReturnSuccess(CompareAction action)
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, action);

        Assert.True(result.Success);
    }

    [Fact]
    public void ComparePresentations_MessageIsNonEmpty()
    {
        var (source, target) = CreateIdenticalPair();

        var result = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.NotEmpty(result.Message);
    }

    [Fact]
    public void ComparePresentations_Idempotent_SameResultTwice()
    {
        var source = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide A" });
        var target = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide B" });

        var result1 = Service.ComparePresentations(source, target, CompareAction.Full);
        var result2 = Service.ComparePresentations(source, target, CompareAction.Full);

        Assert.Equal(result1.AreIdentical, result2.AreIdentical);
        Assert.Equal(result1.DifferenceCount, result2.DifferenceCount);
    }
}
