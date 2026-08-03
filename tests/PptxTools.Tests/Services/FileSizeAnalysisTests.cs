namespace PptxTools.Tests.Services;

/// <summary>
/// Service-level tests for AnalyzeFileSize (Issue #80 — Analyze presentation file size breakdown).
/// Written proactively while Cheritto implements the tool.
/// Validates category breakdown, arithmetic invariants, and edge case handling.
/// </summary>
[Trait("Category", "Unit")]
public class FileSizeAnalysisTests : PptxTestBase
{
    private static readonly string[] ExpectedCategoryNames =
        ["slides", "images", "video_audio", "masters", "layouts", "other"];

    // ────────────────────────────────────────────────────────
    // Happy path: minimal single-slide PPTX
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AnalyzeFileSize_MinimalPptx_ReturnsSuccess()
    {
        var path = CreateMinimalPptx();
        var result = Service.AnalyzeFileSize(path);

        Assert.True(result.Success);
        Assert.Equal(path, result.FilePath);
        Assert.NotEmpty(result.Message);
    }

    [Fact]
    public void AnalyzeFileSize_MinimalPptx_AllSixCategoriesPresent()
    {
        var path = CreateMinimalPptx();
        var result = Service.AnalyzeFileSize(path);

        var categoryNames = result.Categories.Select(c => c.Name).OrderBy(n => n).ToArray();
        var expected = ExpectedCategoryNames.OrderBy(n => n).ToArray();
        Assert.Equal(expected, categoryNames);
    }

    [Fact]
    public void AnalyzeFileSize_MinimalPptx_SlidesCategory_HasOnePart()
    {
        var path = CreateMinimalPptx();
        var result = Service.AnalyzeFileSize(path);

        var slides = Assert.Single(result.Categories, c => c.Name == "slides");
        Assert.Equal(1, slides.PartCount);
        Assert.Single(slides.Parts);
        Assert.True(slides.TotalSize > 0);
    }

    [Fact]
    public void AnalyzeFileSize_MinimalPptx_MastersCategory_HasOnePart()
    {
        var path = CreateMinimalPptx();
        var result = Service.AnalyzeFileSize(path);

        var masters = Assert.Single(result.Categories, c => c.Name == "masters");
        Assert.Equal(1, masters.PartCount);
        Assert.Single(masters.Parts);
        Assert.True(masters.TotalSize > 0);
    }

    [Fact]
    public void AnalyzeFileSize_MinimalPptx_LayoutsCategory_HasOnePart()
    {
        var path = CreateMinimalPptx();
        var result = Service.AnalyzeFileSize(path);

        var layouts = Assert.Single(result.Categories, c => c.Name == "layouts");
        Assert.Equal(1, layouts.PartCount);
        Assert.Single(layouts.Parts);
        Assert.True(layouts.TotalSize > 0);
    }

    // ────────────────────────────────────────────────────────
    // Multi-slide with images
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AnalyzeFileSize_WithImage_ImagesCategory_ContainsImagePart()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide with image", IncludeImage = true });
        var result = Service.AnalyzeFileSize(path);

        var images = Assert.Single(result.Categories, c => c.Name == "images");
        Assert.True(images.PartCount >= 1, "Expected at least 1 image part.");
        Assert.True(images.TotalSize > 0, "Image category should have non-zero size.");
        Assert.Contains(images.Parts, p => p.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AnalyzeFileSize_MultipleSlides_SlidesCategory_MatchesSlideCount()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide 1" },
            new TestSlideDefinition { TitleText = "Slide 2" },
            new TestSlideDefinition { TitleText = "Slide 3", IncludeImage = true });
        var result = Service.AnalyzeFileSize(path);

        var slides = Assert.Single(result.Categories, c => c.Name == "slides");
        Assert.Equal(3, slides.PartCount);
        Assert.Equal(3, slides.Parts.Count);
    }

    [Fact]
    public void AnalyzeFileSize_MultiSlideWithImages_ImageSizeIsPositive()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Img 1", IncludeImage = true },
            new TestSlideDefinition { TitleText = "Img 2", IncludeImage = true });
        var result = Service.AnalyzeFileSize(path);

        var images = Assert.Single(result.Categories, c => c.Name == "images");
        Assert.True(images.PartCount >= 2, "Expected at least 2 image parts for 2 slides with images.");
        Assert.All(images.Parts, p => Assert.True(p.Size > 0, $"Part {p.Path} should have positive size."));
    }

    // ────────────────────────────────────────────────────────
    // Masters / layouts categorized separately
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AnalyzeFileSize_MastersAndLayouts_AreInDistinctCategories()
    {
        var path = CreateMinimalPptx();
        var result = Service.AnalyzeFileSize(path);

        var masters = Assert.Single(result.Categories, c => c.Name == "masters");
        var layouts = Assert.Single(result.Categories, c => c.Name == "layouts");

        // Master and layout paths should not overlap
        var masterPaths = masters.Parts.Select(p => p.Path).ToHashSet();
        var layoutPaths = layouts.Parts.Select(p => p.Path).ToHashSet();
        Assert.Empty(masterPaths.Intersect(layoutPaths));
    }

    [Fact]
    public void AnalyzeFileSize_MasterParts_HaveExpectedPathPrefix()
    {
        var path = CreateMinimalPptx();
        var result = Service.AnalyzeFileSize(path);

        var masters = Assert.Single(result.Categories, c => c.Name == "masters");
        Assert.All(masters.Parts, p =>
            Assert.StartsWith("/ppt/slideMasters/", p.Path, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AnalyzeFileSize_LayoutParts_HaveExpectedPathPrefix()
    {
        var path = CreateMinimalPptx();
        var result = Service.AnalyzeFileSize(path);

        var layouts = Assert.Single(result.Categories, c => c.Name == "layouts");
        Assert.All(layouts.Parts, p =>
            Assert.StartsWith("/ppt/slideLayouts/", p.Path, StringComparison.OrdinalIgnoreCase));
    }

    // ────────────────────────────────────────────────────────
    // Empty / missing categories → 0, not null
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AnalyzeFileSize_EmptyCategories_HaveZeroTotalAndNonNullParts()
    {
        var path = CreateMinimalPptx();
        var result = Service.AnalyzeFileSize(path);

        // A minimal PPTX has no video/audio
        var videoAudio = Assert.Single(result.Categories, c => c.Name == "video_audio");
        Assert.Equal(0, videoAudio.TotalSize);
        Assert.Equal(0, videoAudio.PartCount);
        Assert.NotNull(videoAudio.Parts);
        Assert.Empty(videoAudio.Parts);
    }

    [Fact]
    public void AnalyzeFileSize_MinimalPptx_ImagesCategory_IsEmptyOrSmall()
    {
        // A minimal PPTX (no IncludeImage) should have 0 images
        var path = CreateMinimalPptx();
        var result = Service.AnalyzeFileSize(path);

        var images = Assert.Single(result.Categories, c => c.Name == "images");
        Assert.Equal(0, images.PartCount);
        Assert.NotNull(images.Parts);
        Assert.Empty(images.Parts);
    }

    // ────────────────────────────────────────────────────────
    // Arithmetic invariants
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AnalyzeFileSize_TotalPartSize_EqualsSumOfCategorySubtotals()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Totals check", IncludeImage = true });
        var result = Service.AnalyzeFileSize(path);

        var sumOfSubtotals = result.Categories.Sum(c => c.TotalSize);
        Assert.Equal(sumOfSubtotals, result.TotalPartSize);
    }

    [Fact]
    public void AnalyzeFileSize_CategorySubtotal_EqualsSumOfPartSizes()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Part sum check", IncludeImage = true });
        var result = Service.AnalyzeFileSize(path);

        Assert.All(result.Categories, category =>
        {
            var partSum = category.Parts.Sum(p => p.Size);
            Assert.Equal(partSum, category.TotalSize);
        });
    }

    [Fact]
    public void AnalyzeFileSize_TotalFileSize_MatchesDiskFileSize()
    {
        var path = CreateMinimalPptx();
        var expectedFileSize = new FileInfo(path).Length;
        var result = Service.AnalyzeFileSize(path);

        Assert.Equal(expectedFileSize, result.TotalFileSize);
    }

    [Fact]
    public void AnalyzeFileSize_TotalFileSize_IsPositive()
    {
        var path = CreateMinimalPptx();
        var result = Service.AnalyzeFileSize(path);

        Assert.True(result.TotalFileSize > 0);
        Assert.True(result.TotalPartSize > 0);
    }

    // ────────────────────────────────────────────────────────
    // Part metadata quality
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AnalyzeFileSize_AllParts_HaveNonEmptyPaths()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Metadata check", IncludeImage = true });
        var result = Service.AnalyzeFileSize(path);

        var allParts = result.Categories.SelectMany(c => c.Parts);
        Assert.All(allParts, p => Assert.False(string.IsNullOrWhiteSpace(p.Path), "Part path should not be empty."));
    }

    [Fact]
    public void AnalyzeFileSize_AllParts_HaveNonEmptyContentType()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "ContentType check", IncludeImage = true });
        var result = Service.AnalyzeFileSize(path);

        var allParts = result.Categories.SelectMany(c => c.Parts);
        Assert.All(allParts, p => Assert.False(string.IsNullOrWhiteSpace(p.ContentType), $"Part {p.Path} should have a content type."));
    }

    [Fact]
    public void AnalyzeFileSize_AllParts_HaveNonNegativeSize()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Size check", IncludeImage = true });
        var result = Service.AnalyzeFileSize(path);

        var allParts = result.Categories.SelectMany(c => c.Parts);
        Assert.All(allParts, p => Assert.True(p.Size >= 0, $"Part {p.Path} should not have negative size."));
    }

    // ────────────────────────────────────────────────────────
    // Error handling
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AnalyzeFileSize_FileNotFound_Throws()
    {
        var bogusPath = Path.Join(Path.GetTempPath(), "nonexistent_" + Guid.NewGuid() + ".pptx");

        Assert.ThrowsAny<Exception>(() => Service.AnalyzeFileSize(bogusPath));
    }

    // ────────────────────────────────────────────────────────
    // Complex fixture: tables, charts, and images together
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AnalyzeFileSize_ComplexPresentation_ReportsAllCategories()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Overview",
                TextShapes = [new TestTextShapeDefinition { Paragraphs = ["Bullet 1", "Bullet 2"] }]
            },
            new TestSlideDefinition
            {
                TitleText = "Data Slide",
                IncludeImage = true,
                Tables = [new TestTableDefinition
                {
                    Rows = [["A", "B"], ["1", "2"]]
                }]
            },
            new TestSlideDefinition
            {
                TitleText = "Chart Slide",
                Charts = [new TestChartDefinition
                {
                    ChartType = "Column",
                    Categories = ["Q1", "Q2"],
                    Series = [new TestSeriesDefinition { Name = "Revenue", Values = [100.0, 200.0] }]
                }]
            });

        var result = Service.AnalyzeFileSize(path);

        Assert.True(result.Success);

        // All 6 categories should still be present
        Assert.Equal(ExpectedCategoryNames.Length, result.Categories.Count);

        // Slides: at least 3 (charts/tables may add related slide-level parts)
        var slides = Assert.Single(result.Categories, c => c.Name == "slides");
        Assert.True(slides.PartCount >= 3, $"Expected at least 3 slide parts, got {slides.PartCount}.");

        // Images: at least 1 from the IncludeImage slide
        var images = Assert.Single(result.Categories, c => c.Name == "images");
        Assert.True(images.PartCount >= 1);

        // Other: should capture presentation.xml, .rels, theme, chart XML, etc.
        var other = Assert.Single(result.Categories, c => c.Name == "other");
        Assert.True(other.PartCount > 0, "Other category should have framework parts.");

        // Grand total still checks out
        Assert.Equal(result.Categories.Sum(c => c.TotalSize), result.TotalPartSize);
    }

    [Fact]
    public void AnalyzeFileSize_CategoryPartCount_MatchesPartsListCount()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Consistency", IncludeImage = true });
        var result = Service.AnalyzeFileSize(path);

        Assert.All(result.Categories, category =>
            Assert.Equal(category.PartCount, category.Parts.Count));
    }
}
