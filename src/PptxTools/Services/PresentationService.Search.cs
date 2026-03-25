using System.Text.RegularExpressions;
using PptxTools.Models;

namespace PptxTools.Services;

public partial class PresentationService
{
    /// <summary>Search all slides for shapes containing the specified text.</summary>
    public TextSearchResult SearchText(string filePath, string searchText, bool caseSensitive = false, int? slideNumber = null)
    {
        if (string.IsNullOrEmpty(searchText))
            return new TextSearchResult(false, [], 0, 0, "Search text cannot be empty.");

        var allSlides = GetAllSlideContents(filePath);
        if (slideNumber.HasValue && (slideNumber.Value < 1 || slideNumber.Value > allSlides.Count))
            return new TextSearchResult(false, [], 0, 0, $"Slide {slideNumber.Value} is out of range. The presentation has {allSlides.Count} slide(s).");

        var slides = FilterSlides(allSlides, slideNumber);
        var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var matches = new List<TextSearchMatch>();

        foreach (var slide in slides)
        {
            int slideNum = slide.SlideIndex + 1;
            foreach (var shape in slide.Shapes)
            {
                var text = GetShapeText(shape);
                if (text is null) continue;

                int startIndex = 0;
                while (true)
                {
                    int idx = text.IndexOf(searchText, startIndex, comparison);
                    if (idx < 0) break;

                    matches.Add(new TextSearchMatch(
                        SlideNumber: slideNum,
                        ShapeName: shape.Name,
                        ShapeType: shape.ShapeType,
                        MatchedText: text.Substring(idx, searchText.Length),
                        Context: ExtractContext(text, idx, searchText.Length),
                        MatchIndex: idx));

                    startIndex = idx + 1;
                }
            }
        }

        return new TextSearchResult(true, matches, matches.Count, slides.Count);
    }

    /// <summary>Search all slides for shapes containing text matching a regex pattern.</summary>
    public TextSearchResult SearchByRegex(string filePath, string pattern, bool caseSensitive = false, int? slideNumber = null)
    {
        if (string.IsNullOrEmpty(pattern))
            return new TextSearchResult(false, [], 0, 0, "Regex pattern cannot be empty.");

        var options = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
        Regex regex;
        try
        {
            regex = new Regex(pattern, options, TimeSpan.FromSeconds(5));
        }
        catch (ArgumentException ex)
        {
            return new TextSearchResult(false, [], 0, 0, $"Invalid regex pattern: {ex.Message}");
        }

        var allSlides = GetAllSlideContents(filePath);
        if (slideNumber.HasValue && (slideNumber.Value < 1 || slideNumber.Value > allSlides.Count))
            return new TextSearchResult(false, [], 0, 0, $"Slide {slideNumber.Value} is out of range. The presentation has {allSlides.Count} slide(s).");

        var slides = FilterSlides(allSlides, slideNumber);
        var matches = new List<TextSearchMatch>();

        foreach (var slide in slides)
        {
            int slideNum = slide.SlideIndex + 1;
            foreach (var shape in slide.Shapes)
            {
                var text = GetShapeText(shape);
                if (text is null) continue;

                foreach (Match m in regex.Matches(text))
                {
                    matches.Add(new TextSearchMatch(
                        SlideNumber: slideNum,
                        ShapeName: shape.Name,
                        ShapeType: shape.ShapeType,
                        MatchedText: m.Value,
                        Context: ExtractContext(text, m.Index, m.Length),
                        MatchIndex: m.Index));
                }
            }
        }

        return new TextSearchResult(true, matches, matches.Count, slides.Count);
    }

    /// <summary>Find shapes with no text content across slides.</summary>
    public EmptyShapeResult FindEmptyShapes(string filePath, int? slideNumber = null)
    {
        var allSlides = GetAllSlideContents(filePath);
        if (slideNumber.HasValue && (slideNumber.Value < 1 || slideNumber.Value > allSlides.Count))
            return new EmptyShapeResult(false, [], 0, 0, $"Slide {slideNumber.Value} is out of range. The presentation has {allSlides.Count} slide(s).");

        var slides = FilterSlides(allSlides, slideNumber);
        var empties = new List<EmptyShapeInfo>();

        foreach (var slide in slides)
        {
            int slideNum = slide.SlideIndex + 1;
            foreach (var shape in slide.Shapes)
            {
                if (shape.ShapeType is not ("Text" or "Table")) continue;

                var text = GetShapeText(shape);
                if (string.IsNullOrWhiteSpace(text))
                {
                    empties.Add(new EmptyShapeInfo(
                        SlideNumber: slideNum,
                        ShapeName: shape.Name,
                        ShapeType: shape.ShapeType,
                        IsPlaceholder: shape.IsPlaceholder,
                        PlaceholderType: shape.PlaceholderType));
                }
            }
        }

        return new EmptyShapeResult(true, empties, empties.Count, slides.Count);
    }

    private static IReadOnlyList<SlideContent> FilterSlides(IReadOnlyList<SlideContent> allSlides, int? slideNumber)
    {
        if (slideNumber is null) return allSlides;

        int idx = slideNumber.Value - 1;
        if (idx < 0 || idx >= allSlides.Count)
            return [];

        return [allSlides[idx]];
    }

    /// <summary>Get combined text from a shape, including table cell text.</summary>
    private static string? GetShapeText(ShapeContent shape)
    {
        if (shape.Text is not null) return shape.Text;

        if (shape.TableRows is { Count: > 0 })
        {
            var cells = shape.TableRows.SelectMany(row => row);
            var combined = string.Join(" ", cells.Where(c => !string.IsNullOrEmpty(c)));
            return combined.Length > 0 ? combined : null;
        }

        return null;
    }

    /// <summary>Extract surrounding context around a match position.</summary>
    private static string ExtractContext(string text, int matchStart, int matchLength, int contextChars = 40)
    {
        int start = Math.Max(0, matchStart - contextChars);
        int end = Math.Min(text.Length, matchStart + matchLength + contextChars);

        var context = text[start..end].Replace('\n', ' ').Replace('\r', ' ');

        if (start > 0) context = "..." + context;
        if (end < text.Length) context += "...";

        return context;
    }
}
