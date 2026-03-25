namespace PptxTools.Models;

/// <summary>A single text match found within a shape.</summary>
/// <param name="SlideNumber">1-based slide number where the match was found.</param>
/// <param name="ShapeName">Name of the shape containing the match.</param>
/// <param name="ShapeType">Shape type (Text, Table, etc.).</param>
/// <param name="MatchedText">The actual text that matched the search.</param>
/// <param name="Context">Surrounding text around the match for context.</param>
/// <param name="MatchIndex">Character position of the match within the shape's full text.</param>
public record TextSearchMatch(
    int SlideNumber,
    string ShapeName,
    string ShapeType,
    string MatchedText,
    string Context,
    int MatchIndex);

/// <summary>Aggregated result of a text search across slides.</summary>
/// <param name="Success">Whether the search completed without errors.</param>
/// <param name="Matches">All matches found.</param>
/// <param name="TotalMatches">Total number of matches found.</param>
/// <param name="SlidesSearched">Number of slides that were searched.</param>
/// <param name="Message">Optional message (error or informational).</param>
public record TextSearchResult(
    bool Success,
    IReadOnlyList<TextSearchMatch> Matches,
    int TotalMatches,
    int SlidesSearched,
    string? Message = null);

/// <summary>A shape that contains no text content.</summary>
/// <param name="SlideNumber">1-based slide number where the empty shape was found.</param>
/// <param name="ShapeName">Name of the empty shape.</param>
/// <param name="ShapeType">Shape type (Text, Table, etc.).</param>
/// <param name="IsPlaceholder">Whether the shape is a layout placeholder.</param>
/// <param name="PlaceholderType">Placeholder type if applicable.</param>
public record EmptyShapeInfo(
    int SlideNumber,
    string ShapeName,
    string ShapeType,
    bool IsPlaceholder,
    string? PlaceholderType);

/// <summary>Aggregated result of an empty shape search across slides.</summary>
/// <param name="Success">Whether the search completed without errors.</param>
/// <param name="EmptyShapes">All empty shapes found.</param>
/// <param name="TotalFound">Total number of empty shapes found.</param>
/// <param name="SlidesSearched">Number of slides that were searched.</param>
/// <param name="Message">Optional message (error or informational).</param>
public record EmptyShapeResult(
    bool Success,
    IReadOnlyList<EmptyShapeInfo> EmptyShapes,
    int TotalFound,
    int SlidesSearched,
    string? Message = null);
