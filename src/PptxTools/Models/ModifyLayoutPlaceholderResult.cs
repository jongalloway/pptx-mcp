namespace PptxTools.Models;

/// <summary>Structured result for the ModifyPlaceholder action of pptx_manage_layouts.</summary>
/// <param name="Success">True when the placeholder was updated without errors.</param>
/// <param name="FilePath">Path to the modified presentation file.</param>
/// <param name="LayoutName">Display name of the layout that was updated.</param>
/// <param name="PlaceholderIndex">0-based placeholder ordinal within the layout placeholder sequence.</param>
/// <param name="NewType">The placeholder type value that was set.</param>
/// <param name="NewIdx">The placeholder idx value that was set (if provided).</param>
/// <param name="Message">Human-readable status or error message.</param>
public record ModifyLayoutPlaceholderResult(
    bool Success,
    string FilePath,
    string? LayoutName,
    int? PlaceholderIndex,
    string? NewType,
    int? NewIdx,
    string Message);
