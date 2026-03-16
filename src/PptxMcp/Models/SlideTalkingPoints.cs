namespace PptxMcp.Models;

/// <summary>Key talking points extracted from a single slide.</summary>
/// <param name="SlideIndex">Zero-based index of the slide in the presentation.</param>
/// <param name="Title">Detected title text for the slide, when available.</param>
/// <param name="Points">Highest-ranked talking points for the slide.</param>
public record SlideTalkingPoints(
    int SlideIndex,
    string? Title,
    IReadOnlyList<string> Points);
