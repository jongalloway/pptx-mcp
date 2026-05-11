namespace PptxTools.Models;

public record SlideMasterInfo(
    int Index,
    string Name,
    string? ThemeName,
    int LayoutCount,
    IReadOnlyList<string> LayoutNames,
    int ShapeCount,
    string? BackgroundFill);
