namespace PptxTools.Models;

/// <summary>Actions for the consolidated pptx_compare_presentations tool.</summary>
public enum CompareAction
{
    /// <summary>Run all comparison checks between two presentations.</summary>
    Full,

    /// <summary>Compare only slide-level changes (added, removed).</summary>
    SlidesOnly,

    /// <summary>Compare only text content changes across matching slides.</summary>
    TextOnly,

    /// <summary>Compare only presentation-level metadata.</summary>
    MetadataOnly
}

/// <summary>A slide-level difference between two presentations.</summary>
/// <param name="SlideNumber">1-based slide number where the difference was detected.</param>
/// <param name="DifferenceType">Classification: Added, Removed, or Modified.</param>
/// <param name="Description">Human-readable description of the difference.</param>
public record SlideDifference(
    int SlideNumber,
    string DifferenceType,
    string? Description = null);

/// <summary>A text difference detected within a shape on a slide.</summary>
/// <param name="SlideNumber">1-based slide number where the text difference was found.</param>
/// <param name="ShapeName">Name of the shape containing the changed text.</param>
/// <param name="DifferenceType">Classification: Added, Removed, or Modified.</param>
/// <param name="SourceText">Text in the source presentation. Null for added shapes.</param>
/// <param name="TargetText">Text in the target presentation. Null for removed shapes.</param>
public record TextDifference(
    int SlideNumber,
    string ShapeName,
    string DifferenceType,
    string? SourceText,
    string? TargetText);

/// <summary>A metadata field that differs between the two presentations.</summary>
/// <param name="Property">Name of the metadata property (e.g. Title, Creator).</param>
/// <param name="SourceValue">Value in the source presentation.</param>
/// <param name="TargetValue">Value in the target presentation.</param>
public record MetadataDifference(
    string Property,
    string? SourceValue,
    string? TargetValue);

/// <summary>Full comparison result between two presentations.</summary>
/// <param name="Success">Whether the comparison completed without errors.</param>
/// <param name="Action">The comparison action that was performed.</param>
/// <param name="SourceFile">Path to the source presentation.</param>
/// <param name="TargetFile">Path to the target presentation.</param>
/// <param name="AreIdentical">True when no differences were found.</param>
/// <param name="DifferenceCount">Total number of differences across all categories.</param>
/// <param name="SlideDifferences">Slide-level differences (added/removed slides).</param>
/// <param name="TextDifferences">Text-level differences across matching slides.</param>
/// <param name="MetadataDifferences">Metadata property differences.</param>
/// <param name="Message">Human-readable summary or error message.</param>
public record ComparisonResult(
    bool Success,
    string Action,
    string SourceFile,
    string TargetFile,
    bool AreIdentical,
    int DifferenceCount,
    IReadOnlyList<SlideDifference>? SlideDifferences,
    IReadOnlyList<TextDifference>? TextDifferences,
    IReadOnlyList<MetadataDifference>? MetadataDifferences,
    string Message);
