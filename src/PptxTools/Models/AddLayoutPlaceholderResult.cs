namespace PptxTools.Models;

/// <summary>Structured result for the AddPlaceholder action of pptx_manage_layouts.</summary>
/// <param name="Success">True when the placeholder was added without errors.</param>
/// <param name="FilePath">Path to the modified presentation file.</param>
/// <param name="LayoutName">Display name of the layout that was updated.</param>
/// <param name="Type">The placeholder type value that was added.</param>
/// <param name="Idx">The placeholder idx value that was added.</param>
/// <param name="X">Placeholder X position in EMU.</param>
/// <param name="Y">Placeholder Y position in EMU.</param>
/// <param name="Cx">Placeholder width in EMU.</param>
/// <param name="Cy">Placeholder height in EMU.</param>
/// <param name="Message">Human-readable status or error message.</param>
public record AddLayoutPlaceholderResult(
    bool Success,
    string FilePath,
    string? LayoutName,
    string? Type,
    int? Idx,
    long? X,
    long? Y,
    long? Cx,
    long? Cy,
    string Message);
