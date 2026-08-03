namespace PptxTools.Models;

/// <summary>Structured result for the SetType action of pptx_manage_layouts.</summary>
/// <param name="Success">True when the type was updated without errors.</param>
/// <param name="FilePath">Path to the modified presentation file.</param>
/// <param name="LayoutName">Display name of the layout that was updated.</param>
/// <param name="LayoutType">The type value that was set on the layout.</param>
/// <param name="Message">Human-readable status or error message.</param>
public record SetLayoutTypeResult(
    bool Success,
    string FilePath,
    string? LayoutName,
    string? LayoutType,
    string Message);
