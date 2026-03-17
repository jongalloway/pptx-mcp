namespace PptxMcp.Tests.Services;

[Trait("Category", "Unit")]
public class SlideOrganizationTests : PptxTestBase
{
    private string CreateNamedSlides(params string[] titles) =>
        CreatePptxWithSlides(titles.Select(t => new TestSlideDefinition { TitleText = t }).ToArray());

    private string[] GetSlideTitles(string path) =>
        Service.GetSlides(path).Select(s => s.Title ?? string.Empty).ToArray();

    // ── MoveSlide ───────────────────────────────────────────────────────────

    [Fact]
    public void MoveSlide_MovesSlideToNewPosition()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.MoveSlide(path, 1, 3);
        Assert.Equal(["B", "C", "A"], GetSlideTitles(path));
    }

    [Fact]
    public void MoveSlide_MoveLastSlideToFirst()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.MoveSlide(path, 3, 1);
        Assert.Equal(["C", "A", "B"], GetSlideTitles(path));
    }

    [Fact]
    public void MoveSlide_MoveMiddleToFirst()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.MoveSlide(path, 2, 1);
        Assert.Equal(["B", "A", "C"], GetSlideTitles(path));
    }

    [Fact]
    public void MoveSlide_SamePositionIsNoOp()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.MoveSlide(path, 2, 2);
        Assert.Equal(["A", "B", "C"], GetSlideTitles(path));
    }

    [Fact]
    public void MoveSlide_PreservesSlideCount()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.MoveSlide(path, 1, 3);
        Assert.Equal(3, Service.GetSlides(path).Count);
    }

    [Fact]
    public void MoveSlide_ThrowsOnInvalidSlideNumber()
    {
        var path = CreateNamedSlides("A", "B");
        Assert.Throws<ArgumentOutOfRangeException>(() => Service.MoveSlide(path, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => Service.MoveSlide(path, 3, 1));
    }

    [Fact]
    public void MoveSlide_ThrowsOnInvalidTargetPosition()
    {
        var path = CreateNamedSlides("A", "B");
        Assert.Throws<ArgumentOutOfRangeException>(() => Service.MoveSlide(path, 1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => Service.MoveSlide(path, 1, 3));
    }

    // ── DeleteSlide ─────────────────────────────────────────────────────────

    [Fact]
    public void DeleteSlide_RemovesSlideFromPresentation()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.DeleteSlide(path, 2);
        Assert.Equal(["A", "C"], GetSlideTitles(path));
    }

    [Fact]
    public void DeleteSlide_DeleteFirstSlide()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.DeleteSlide(path, 1);
        Assert.Equal(["B", "C"], GetSlideTitles(path));
    }

    [Fact]
    public void DeleteSlide_DeleteLastSlide()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.DeleteSlide(path, 3);
        Assert.Equal(["A", "B"], GetSlideTitles(path));
    }

    [Fact]
    public void DeleteSlide_DecreasesSlideCount()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.DeleteSlide(path, 1);
        Assert.Equal(2, Service.GetSlides(path).Count);
    }

    [Fact]
    public void DeleteSlide_ThrowsWhenOnlyOneSlide()
    {
        var path = CreateNamedSlides("Only Slide");
        Assert.Throws<InvalidOperationException>(() => Service.DeleteSlide(path, 1));
    }

    [Fact]
    public void DeleteSlide_ThrowsOnInvalidSlideNumber()
    {
        var path = CreateNamedSlides("A", "B");
        Assert.Throws<ArgumentOutOfRangeException>(() => Service.DeleteSlide(path, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => Service.DeleteSlide(path, 3));
    }

    // ── ReorderSlides ────────────────────────────────────────────────────────

    [Fact]
    public void ReorderSlides_ReversesOrder()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.ReorderSlides(path, [3, 2, 1]);
        Assert.Equal(["C", "B", "A"], GetSlideTitles(path));
    }

    [Fact]
    public void ReorderSlides_IdentityOrderIsNoOp()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.ReorderSlides(path, [1, 2, 3]);
        Assert.Equal(["A", "B", "C"], GetSlideTitles(path));
    }

    [Fact]
    public void ReorderSlides_ArbitraryPermutation()
    {
        var path = CreateNamedSlides("A", "B", "C", "D");
        Service.ReorderSlides(path, [3, 1, 4, 2]);
        Assert.Equal(["C", "A", "D", "B"], GetSlideTitles(path));
    }

    [Fact]
    public void ReorderSlides_PreservesSlideCount()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Service.ReorderSlides(path, [2, 3, 1]);
        Assert.Equal(3, Service.GetSlides(path).Count);
    }

    [Fact]
    public void ReorderSlides_ThrowsWhenLengthMismatch()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Assert.Throws<ArgumentException>(() => Service.ReorderSlides(path, [1, 2]));
        Assert.Throws<ArgumentException>(() => Service.ReorderSlides(path, [1, 2, 3, 4]));
    }

    [Fact]
    public void ReorderSlides_ThrowsWhenNotAPermutation()
    {
        var path = CreateNamedSlides("A", "B", "C");
        Assert.Throws<ArgumentException>(() => Service.ReorderSlides(path, [1, 1, 3]));
        Assert.Throws<ArgumentException>(() => Service.ReorderSlides(path, [1, 2, 4]));
    }

    // ── Relationship Integrity ───────────────────────────────────────────────

    [Fact]
    public void MoveSlide_ContentIsPreservedAfterMove()
    {
        var path = CreateNamedSlides("Intro", "Middle", "Conclusion");
        Service.MoveSlide(path, 3, 1);
        var slides = Service.GetSlides(path);
        Assert.Equal("Conclusion", slides[0].Title);
        Assert.Equal("Intro", slides[1].Title);
        Assert.Equal("Middle", slides[2].Title);
    }

    [Fact]
    public void DeleteSlide_RemainingSlideContentIsIntact()
    {
        var path = CreateNamedSlides("Keep1", "Delete", "Keep2");
        Service.DeleteSlide(path, 2);
        var slides = Service.GetSlides(path);
        Assert.Equal("Keep1", slides[0].Title);
        Assert.Equal("Keep2", slides[1].Title);
    }

    [Fact]
    public void ReorderSlides_SlideContentIsPreservedAfterReorder()
    {
        var path = CreateNamedSlides("First", "Second", "Third");
        Service.ReorderSlides(path, [3, 1, 2]);
        var slides = Service.GetSlides(path);
        Assert.Equal("Third", slides[0].Title);
        Assert.Equal("First", slides[1].Title);
        Assert.Equal("Second", slides[2].Title);
    }
}
