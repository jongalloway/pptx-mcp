namespace PptxTools.Models;

/// <summary>Color scheme extracted from a presentation theme.</summary>
/// <param name="Name">Name of the color scheme.</param>
/// <param name="Dark1">Dark1 color in #RRGGBB format.</param>
/// <param name="Light1">Light1 color in #RRGGBB format.</param>
/// <param name="Dark2">Dark2 color in #RRGGBB format.</param>
/// <param name="Light2">Light2 color in #RRGGBB format.</param>
/// <param name="Accent1">Accent1 color in #RRGGBB format.</param>
/// <param name="Accent2">Accent2 color in #RRGGBB format.</param>
/// <param name="Accent3">Accent3 color in #RRGGBB format.</param>
/// <param name="Accent4">Accent4 color in #RRGGBB format.</param>
/// <param name="Accent5">Accent5 color in #RRGGBB format.</param>
/// <param name="Accent6">Accent6 color in #RRGGBB format.</param>
/// <param name="Hyperlink">Hyperlink (followed link) color in #RRGGBB format.</param>
/// <param name="FollowedHyperlink">Followed hyperlink color in #RRGGBB format.</param>
public record ThemeColorScheme(
    string? Name,
    string? Dark1,
    string? Light1,
    string? Dark2,
    string? Light2,
    string? Accent1,
    string? Accent2,
    string? Accent3,
    string? Accent4,
    string? Accent5,
    string? Accent6,
    string? Hyperlink,
    string? FollowedHyperlink);

/// <summary>Font scheme extracted from a presentation theme.</summary>
/// <param name="Name">Name of the font scheme.</param>
/// <param name="MajorFont">Major (heading) font typeface.</param>
/// <param name="MinorFont">Minor (body) font typeface.</param>
public record ThemeFontScheme(
    string? Name,
    string? MajorFont,
    string? MinorFont);

/// <summary>Theme information for a single slide master.</summary>
/// <param name="MasterIndex">0-based index of the slide master this theme belongs to.</param>
/// <param name="ThemeName">Name of the theme.</param>
/// <param name="ColorScheme">The resolved color scheme.</param>
/// <param name="FontScheme">The font scheme.</param>
public record ThemeInfo(
    int MasterIndex,
    string? ThemeName,
    ThemeColorScheme? ColorScheme,
    ThemeFontScheme? FontScheme);

/// <summary>Result returned by the pptx_inspect_theme tool.</summary>
/// <param name="Success">Whether the operation succeeded.</param>
/// <param name="ThemeCount">Number of themes found (one per slide master).</param>
/// <param name="Themes">Theme details for each slide master.</param>
/// <param name="Message">Human-readable status or error message.</param>
public record ThemeInfoResult(
    bool Success,
    int ThemeCount,
    IReadOnlyList<ThemeInfo> Themes,
    string Message);
