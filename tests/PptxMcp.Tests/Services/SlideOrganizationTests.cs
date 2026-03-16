namespace PptxMcp.Tests.Services;

public class SlideOrganizationTests : IDisposable
{
    private readonly PresentationService _service = new();
    private readonly List<string> _tempFiles = new();

    public void Dispose()
    {
        foreach (var file in _tempFiles)
            if (File.Exists(file)) File.Delete(file);
    }

    private string CreatePptxWithSlides(params string[] titles)
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        _tempFiles.Add(path);
        var slides = titles.Select(t => new TestSlideDefinition { TitleText = t }).ToArray();
        TestPptxHelper.CreatePresentation(path, slides);
        return path;
    }

    private string[] GetSlideTitles(string path) =>
        _service.GetSlides(path).Select(s => s.Title ?? string.Empty).ToArray();

    // ── MoveSlide ───────────────────────────────────────────────────────────

    [Fact]
    public void MoveSlide_MovesSlideToNewPosition()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.MoveSlide(path, 1, 3);
        Assert.Equal(["B", "C", "A"], GetSlideTitles(path));
    }

    [Fact]
    public void MoveSlide_MoveLastSlideToFirst()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.MoveSlide(path, 3, 1);
        Assert.Equal(["C", "A", "B"], GetSlideTitles(path));
    }

    [Fact]
    public void MoveSlide_MoveMiddleToFirst()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.MoveSlide(path, 2, 1);
        Assert.Equal(["B", "A", "C"], GetSlideTitles(path));
    }

    [Fact]
    public void MoveSlide_SamePositionIsNoOp()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.MoveSlide(path, 2, 2);
        Assert.Equal(["A", "B", "C"], GetSlideTitles(path));
    }

    [Fact]
    public void MoveSlide_PreservesSlideCount()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.MoveSlide(path, 1, 3);
        Assert.Equal(3, _service.GetSlides(path).Count);
    }

    [Fact]
    public void MoveSlide_ThrowsOnInvalidSlideNumber()
    {
        var path = CreatePptxWithSlides("A", "B");
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.MoveSlide(path, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.MoveSlide(path, 3, 1));
    }

    [Fact]
    public void MoveSlide_ThrowsOnInvalidTargetPosition()
    {
        var path = CreatePptxWithSlides("A", "B");
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.MoveSlide(path, 1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.MoveSlide(path, 1, 3));
    }

    // ── DeleteSlide ─────────────────────────────────────────────────────────

    [Fact]
    public void DeleteSlide_RemovesSlideFromPresentation()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.DeleteSlide(path, 2);
        Assert.Equal(["A", "C"], GetSlideTitles(path));
    }

    [Fact]
    public void DeleteSlide_DeleteFirstSlide()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.DeleteSlide(path, 1);
        Assert.Equal(["B", "C"], GetSlideTitles(path));
    }

    [Fact]
    public void DeleteSlide_DeleteLastSlide()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.DeleteSlide(path, 3);
        Assert.Equal(["A", "B"], GetSlideTitles(path));
    }

    [Fact]
    public void DeleteSlide_DecreasesSlideCount()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.DeleteSlide(path, 1);
        Assert.Equal(2, _service.GetSlides(path).Count);
    }

    [Fact]
    public void DeleteSlide_ThrowsWhenOnlyOneSlide()
    {
        var path = CreatePptxWithSlides("Only Slide");
        Assert.Throws<InvalidOperationException>(() => _service.DeleteSlide(path, 1));
    }

    [Fact]
    public void DeleteSlide_ThrowsOnInvalidSlideNumber()
    {
        var path = CreatePptxWithSlides("A", "B");
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.DeleteSlide(path, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.DeleteSlide(path, 3));
    }

    // ── ReorderSlides ────────────────────────────────────────────────────────

    [Fact]
    public void ReorderSlides_ReversesOrder()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.ReorderSlides(path, [3, 2, 1]);
        Assert.Equal(["C", "B", "A"], GetSlideTitles(path));
    }

    [Fact]
    public void ReorderSlides_IdentityOrderIsNoOp()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.ReorderSlides(path, [1, 2, 3]);
        Assert.Equal(["A", "B", "C"], GetSlideTitles(path));
    }

    [Fact]
    public void ReorderSlides_ArbitraryPermutation()
    {
        var path = CreatePptxWithSlides("A", "B", "C", "D");
        _service.ReorderSlides(path, [3, 1, 4, 2]);
        Assert.Equal(["C", "A", "D", "B"], GetSlideTitles(path));
    }

    [Fact]
    public void ReorderSlides_PreservesSlideCount()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        _service.ReorderSlides(path, [2, 3, 1]);
        Assert.Equal(3, _service.GetSlides(path).Count);
    }

    [Fact]
    public void ReorderSlides_ThrowsWhenLengthMismatch()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        Assert.Throws<ArgumentException>(() => _service.ReorderSlides(path, [1, 2]));
        Assert.Throws<ArgumentException>(() => _service.ReorderSlides(path, [1, 2, 3, 4]));
    }

    [Fact]
    public void ReorderSlides_ThrowsWhenNotAPermutation()
    {
        var path = CreatePptxWithSlides("A", "B", "C");
        Assert.Throws<ArgumentException>(() => _service.ReorderSlides(path, [1, 1, 3]));
        Assert.Throws<ArgumentException>(() => _service.ReorderSlides(path, [1, 2, 4]));
    }

    // ── Relationship Integrity ───────────────────────────────────────────────

    [Fact]
    public void MoveSlide_ContentIsPreservedAfterMove()
    {
        var path = CreatePptxWithSlides("Intro", "Middle", "Conclusion");
        _service.MoveSlide(path, 3, 1);
        var slides = _service.GetSlides(path);
        Assert.Equal("Conclusion", slides[0].Title);
        Assert.Equal("Intro", slides[1].Title);
        Assert.Equal("Middle", slides[2].Title);
    }

    [Fact]
    public void DeleteSlide_RemainingSlideContentIsIntact()
    {
        var path = CreatePptxWithSlides("Keep1", "Delete", "Keep2");
        _service.DeleteSlide(path, 2);
        var slides = _service.GetSlides(path);
        Assert.Equal("Keep1", slides[0].Title);
        Assert.Equal("Keep2", slides[1].Title);
    }

    [Fact]
    public void ReorderSlides_SlideContentIsPreservedAfterReorder()
    {
        var path = CreatePptxWithSlides("First", "Second", "Third");
        _service.ReorderSlides(path, [3, 1, 2]);
        var slides = _service.GetSlides(path);
        Assert.Equal("Third", slides[0].Title);
        Assert.Equal("First", slides[1].Title);
        Assert.Equal("Second", slides[2].Title);
    }
}
