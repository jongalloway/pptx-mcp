namespace PptxTools.Tests.Services;

[Trait("Category", "Unit")]
public class MediaAnalysisTests : PptxTestBase
{
    // ── No media ───────────────────────────────────────────────────────

    [Fact]
    public void AnalyzeMedia_NoMedia_ReturnsSuccessWithZeroCounts()
    {
        var path = CreateMinimalPptx();

        var result = Service.AnalyzeMedia(path);

        Assert.True(result.Success);
        Assert.Equal(0, result.TotalMediaCount);
        Assert.Equal(0L, result.TotalMediaSize);
        Assert.Equal(0, result.DuplicateGroupCount);
        Assert.Equal(0L, result.DuplicateSavingsBytes);
    }

    [Fact]
    public void AnalyzeMedia_NoMedia_ReturnsEmptyListsNotNull()
    {
        var path = CreateMinimalPptx();

        var result = Service.AnalyzeMedia(path);

        Assert.NotNull(result.MediaParts);
        Assert.Empty(result.MediaParts);
        Assert.NotNull(result.DuplicateGroups);
        Assert.Empty(result.DuplicateGroups);
    }

    [Fact]
    public void AnalyzeMedia_NoMedia_ReturnsFilePath()
    {
        var path = CreateMinimalPptx();

        var result = Service.AnalyzeMedia(path);

        Assert.Equal(path, result.FilePath);
    }

    [Fact]
    public void AnalyzeMedia_NoMedia_MessageIndicatesNoAssets()
    {
        var path = CreateMinimalPptx();

        var result = Service.AnalyzeMedia(path);

        Assert.NotNull(result.Message);
        Assert.NotEmpty(result.Message);
    }

    // ── Single image ───────────────────────────────────────────────────

    [Fact]
    public void AnalyzeMedia_SingleImage_ReturnsAtLeastOneMediaPart()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.True(result.Success);
        Assert.True(result.TotalMediaCount >= 1);
        Assert.NotEmpty(result.MediaParts);
    }

    [Fact]
    public void AnalyzeMedia_SingleImage_PartHasPositiveSize()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.All(result.MediaParts, part => Assert.True(part.SizeBytes > 0));
    }

    [Fact]
    public void AnalyzeMedia_SingleImage_PartHasNonEmptyHash()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.All(result.MediaParts, part =>
        {
            Assert.NotNull(part.Hash);
            Assert.NotEmpty(part.Hash);
        });
    }

    [Fact]
    public void AnalyzeMedia_SingleImage_TotalMediaSizeIsPositive()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.True(result.TotalMediaSize > 0);
    }

    // ── Multiple images ────────────────────────────────────────────────

    [Fact]
    public void AnalyzeMedia_MultipleImages_ReturnsAtLeastTwoParts()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.True(result.Success);
        Assert.True(result.TotalMediaCount >= 2);
    }

    [Fact]
    public void AnalyzeMedia_MultipleImages_AllPartsHaveHashes()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.All(result.MediaParts, part =>
        {
            Assert.NotNull(part.Hash);
            Assert.NotEmpty(part.Hash);
        });
    }

    // ── Duplicate detection ────────────────────────────────────────────
    // The test helper uses the same SampleImageBytes for every IncludeImage slide.
    // Each slide gets its own ImagePart, so they are separate package parts with
    // identical content → the analyzer should group them as duplicates.

    [Fact]
    public void AnalyzeMedia_DuplicateImages_DetectsDuplicateGroup()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.True(result.DuplicateGroupCount >= 1);
        Assert.NotEmpty(result.DuplicateGroups);
    }

    [Fact]
    public void AnalyzeMedia_DuplicateImages_GroupContainsMultipleParts()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        var group = Assert.Single(result.DuplicateGroups,
            g => g.Parts.Length >= 2);
        Assert.True(group.Parts.Length >= 2);
    }

    [Fact]
    public void AnalyzeMedia_DuplicateImages_SavingsBytesIsPositive()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.True(result.DuplicateSavingsBytes > 0);
    }

    [Fact]
    public void AnalyzeMedia_DuplicateImages_GroupHashMatchesPartHashes()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        foreach (var group in result.DuplicateGroups)
        {
            var partsInGroup = result.MediaParts
                .Where(p => group.Parts.Contains(p.Path))
                .ToList();

            Assert.All(partsInGroup, p => Assert.Equal(group.Hash, p.Hash));
        }
    }

    // ── Hash consistency ───────────────────────────────────────────────

    [Fact]
    public void AnalyzeMedia_SameFile_ProducesSameHashesAcrossCalls()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result1 = Service.AnalyzeMedia(path);
        var result2 = Service.AnalyzeMedia(path);

        Assert.Equal(result1.MediaParts.Length, result2.MediaParts.Length);
        for (int i = 0; i < result1.MediaParts.Length; i++)
        {
            Assert.Equal(result1.MediaParts[i].Hash, result2.MediaParts[i].Hash);
        }
    }

    // ── Arithmetic invariants ──────────────────────────────────────────

    [Fact]
    public void AnalyzeMedia_TotalMediaSizeEqualsSumOfPartSizes()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        var expectedTotal = result.MediaParts.Sum(p => p.SizeBytes);
        Assert.Equal(expectedTotal, result.TotalMediaSize);
    }

    [Fact]
    public void AnalyzeMedia_DuplicateSavingsBytes_IsNonNegative()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.True(result.DuplicateSavingsBytes >= 0);
    }

    [Fact]
    public void AnalyzeMedia_DuplicateSavingsBytes_EqualsExpectedFormula()
    {
        // Savings = sum over each group of: SizeBytes * (Parts.Length - 1)
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        var expectedSavings = result.DuplicateGroups
            .Sum(g => g.SizeBytes * (g.Parts.Length - 1));
        Assert.Equal(expectedSavings, result.DuplicateSavingsBytes);
    }

    [Fact]
    public void AnalyzeMedia_TotalMediaCountMatchesMediaPartsLength()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.Equal(result.MediaParts.Length, result.TotalMediaCount);
    }

    [Fact]
    public void AnalyzeMedia_DuplicateGroupCountMatchesDuplicateGroupsLength()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.Equal(result.DuplicateGroups.Length, result.DuplicateGroupCount);
    }

    // ── Metadata quality ───────────────────────────────────────────────

    [Fact]
    public void AnalyzeMedia_AllPartsHaveNonEmptyPath()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.All(result.MediaParts, part =>
        {
            Assert.NotNull(part.Path);
            Assert.NotEmpty(part.Path);
        });
    }

    [Fact]
    public void AnalyzeMedia_AllPartsHaveNonEmptyContentType()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.All(result.MediaParts, part =>
        {
            Assert.NotNull(part.ContentType);
            Assert.NotEmpty(part.ContentType);
        });
    }

    [Fact]
    public void AnalyzeMedia_ImagePartContentTypeContainsImage()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.Contains(result.MediaParts, part =>
            part.ContentType.Contains("image", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AnalyzeMedia_AllPartSizesAreNonNegative()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.All(result.MediaParts, part => Assert.True(part.SizeBytes >= 0));
    }

    // ── ReferencedBySlides ─────────────────────────────────────────────

    [Fact]
    public void AnalyzeMedia_SingleSlideImage_ReferencedBySlideOne()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        var slideImageParts = result.MediaParts
            .Where(p => p.ReferencedBySlides.Contains(1))
            .ToList();

        Assert.NotEmpty(slideImageParts);
    }

    [Fact]
    public void AnalyzeMedia_ReferencedBySlides_AllValuesArePositive()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        foreach (var part in result.MediaParts)
        {
            Assert.All(part.ReferencedBySlides,
                slideNum => Assert.True(slideNum >= 1,
                    $"Slide number {slideNum} should be >= 1 (1-based)"));
        }
    }

    [Fact]
    public void AnalyzeMedia_TwoSlidesWithImages_ReferencedByCorrectSlideNumbers()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        var allReferencedSlides = result.MediaParts
            .SelectMany(p => p.ReferencedBySlides)
            .Where(n => n > 0)
            .Distinct()
            .OrderBy(n => n)
            .ToArray();

        Assert.Contains(1, allReferencedSlides);
        Assert.Contains(2, allReferencedSlides);
    }

    [Fact]
    public void AnalyzeMedia_DuplicateGroupReferencedBySlides_ContainsAllSlides()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        foreach (var group in result.DuplicateGroups)
        {
            var expectedSlides = result.MediaParts
                .Where(p => group.Parts.Contains(p.Path))
                .SelectMany(p => p.ReferencedBySlides)
                .Distinct()
                .OrderBy(n => n)
                .ToArray();

            Assert.Equal(expectedSlides, group.ReferencedBySlides);
        }
    }

    // ── File not found ─────────────────────────────────────────────────

    [Fact]
    public void AnalyzeMedia_FileNotFound_ThrowsException()
    {
        var nonExistentPath = Path.Join(Path.GetTempPath(), "nonexistent_" + Guid.NewGuid() + ".pptx");

        Assert.ThrowsAny<Exception>(() => Service.AnalyzeMedia(nonExistentPath));
    }

    // ── Mixed slides (with and without images) ─────────────────────────

    [Fact]
    public void AnalyzeMedia_MixedSlides_OnlyCountsMediaFromImageSlides()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Text Only" },
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { TitleText = "Another Text" });

        var result = Service.AnalyzeMedia(path);

        Assert.True(result.Success);
        Assert.True(result.TotalMediaCount >= 1);
    }

    [Fact]
    public void AnalyzeMedia_NoDuplicates_DuplicateGroupCountIsZero()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.Equal(0, result.DuplicateGroupCount);
        Assert.Empty(result.DuplicateGroups);
        Assert.Equal(0L, result.DuplicateSavingsBytes);
    }

    // ── Message content ────────────────────────────────────────────────

    [Fact]
    public void AnalyzeMedia_WithMedia_MessageContainsAssetCount()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        Assert.Contains("media asset", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AnalyzeMedia_WithDuplicates_MessageContainsDuplicateInfo()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        if (result.DuplicateGroupCount > 0)
        {
            Assert.Contains("duplicate", result.Message, StringComparison.OrdinalIgnoreCase);
        }
    }

    // ── MediaParts ordering ────────────────────────────────────────────

    [Fact]
    public void AnalyzeMedia_MediaPartsAreOrderedByPath()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true },
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeMedia(path);

        var paths = result.MediaParts.Select(p => p.Path).ToArray();
        var sortedPaths = paths.OrderBy(p => p).ToArray();
        Assert.Equal(sortedPaths, paths);
    }
}
