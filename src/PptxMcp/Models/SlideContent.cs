namespace PptxMcp.Models;

/// <summary>Structured content of a single slide, including slide dimensions and all shapes.</summary>
/// <param name="SlideIndex">Zero-based index of this slide in the presentation.</param>
/// <param name="SlideWidthEmu">Slide width in EMUs as declared in the presentation (typically 9144000 for 10 inches).</param>
/// <param name="SlideHeightEmu">Slide height in EMUs (typically 6858000 for 7.5 inches, or 5143500 for 4:3).</param>
/// <param name="Shapes">All shapes found on the slide's shape tree, in document order.</param>
public record SlideContent(
    int SlideIndex,
    long SlideWidthEmu,
    long SlideHeightEmu,
    IReadOnlyList<ShapeContent> Shapes);
