namespace PptxTools.Models;

/// <summary>Actions for the consolidated pptx_manage_text_formatting tool.</summary>
public enum TextFormattingAction
{
    /// <summary>Read text formatting properties from shapes on a slide.</summary>
    Get,

    /// <summary>Apply font styling to all runs in a target shape.</summary>
    Apply
}

/// <summary>Formatting details for a single text run within a shape.</summary>
/// <param name="SlideNumber">1-based slide number.</param>
/// <param name="ShapeName">Name of the shape containing this text.</param>
/// <param name="Text">Text snippet from this run.</param>
/// <param name="FontFamily">Latin font typeface, if set.</param>
/// <param name="FontSize">Font size in points, if set.</param>
/// <param name="Bold">Whether the run is bold, if set.</param>
/// <param name="Italic">Whether the run is italic, if set.</param>
/// <param name="Underline">Whether the run is underlined, if set.</param>
/// <param name="Color">Hex RGB color (e.g. "#FF0000"), if set.</param>
/// <param name="Alignment">Paragraph alignment (Left, Center, Right, Justify), if set.</param>
public record TextFormattingInfo(
    int SlideNumber,
    string ShapeName,
    string? Text,
    string? FontFamily,
    double? FontSize,
    bool? Bold,
    bool? Italic,
    bool? Underline,
    string? Color,
    string? Alignment);

/// <summary>Result returned by the pptx_manage_text_formatting tool.</summary>
/// <param name="Success">Whether the operation succeeded.</param>
/// <param name="Action">The action that was performed.</param>
/// <param name="SlideNumber">Slide number filter/target (if applicable).</param>
/// <param name="ShapeName">Shape name filter/target (if applicable).</param>
/// <param name="FormattingCount">Number of formatting entries returned or runs modified.</param>
/// <param name="Formattings">List of formatting details (populated for Get action).</param>
/// <param name="Message">Human-readable status or error message.</param>
public record TextFormattingResult(
    bool Success,
    string Action,
    int? SlideNumber,
    string? ShapeName,
    int FormattingCount,
    IReadOnlyList<TextFormattingInfo> Formattings,
    string Message);
