namespace PptxTools.Models;

/// <summary>Actions for the consolidated pptx_manage_hyperlinks tool.</summary>
public enum HyperlinkAction
{
    /// <summary>Get all hyperlinks in the presentation, optionally filtered by slide number.</summary>
    Get,

    /// <summary>Add an external hyperlink to a shape.</summary>
    Add,

    /// <summary>Update the URL of an existing hyperlink on a shape.</summary>
    Update,

    /// <summary>Remove a hyperlink from a shape.</summary>
    Remove
}

/// <summary>A hyperlink found in the presentation.</summary>
/// <param name="SlideNumber">1-based slide number where the hyperlink was found.</param>
/// <param name="ShapeName">Name of the shape containing the hyperlink.</param>
/// <param name="Text">Display text of the hyperlinked run, if text-level. Null for shape-level hyperlinks.</param>
/// <param name="Url">External URL target. Null for internal slide links.</param>
/// <param name="TargetSlideNumber">1-based target slide number for internal links. Null for external links.</param>
/// <param name="Tooltip">Optional tooltip text shown on hover.</param>
/// <param name="HyperlinkType">Classification: "external", "internal", or "email".</param>
public record HyperlinkInfo(
    int SlideNumber,
    string ShapeName,
    string? Text,
    string? Url,
    int? TargetSlideNumber,
    string? Tooltip,
    string HyperlinkType);

/// <summary>Result of a hyperlink management operation.</summary>
public record HyperlinkResult(
    bool Success,
    string Action,
    int? SlideNumber,
    string? ShapeName,
    string? Url,
    int HyperlinkCount,
    IReadOnlyList<HyperlinkInfo>? Hyperlinks,
    string Message);
