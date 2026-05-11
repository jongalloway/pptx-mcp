using DocumentFormat.OpenXml.Packaging;
using PptxTools.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxTools.Services;

public partial class PresentationService
{
    /// <summary>
    /// Extracts theme color and font information from every slide master in the presentation.
    /// </summary>
    public ThemeInfoResult GetThemeInfo(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("The file does not contain a presentation part.");

        var themes = new List<ThemeInfo>();
        int masterIndex = 0;

        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            var themePart = masterPart.ThemePart;
            if (themePart is null)
            {
                themes.Add(new ThemeInfo(masterIndex, null, null, null));
                masterIndex++;
                continue;
            }

            var theme = themePart.Theme;
            var themeElements = theme?.ThemeElements;

            string? themeName = theme?.Name?.Value;

            ThemeColorScheme? colorScheme = null;
            var clrScheme = themeElements?.ColorScheme;
            if (clrScheme is not null)
            {
                colorScheme = new ThemeColorScheme(
                    Name: clrScheme.Name?.Value,
                    Dark1: ResolveColor(clrScheme.Dark1Color),
                    Light1: ResolveColor(clrScheme.Light1Color),
                    Dark2: ResolveColor(clrScheme.Dark2Color),
                    Light2: ResolveColor(clrScheme.Light2Color),
                    Accent1: ResolveColor(clrScheme.Accent1Color),
                    Accent2: ResolveColor(clrScheme.Accent2Color),
                    Accent3: ResolveColor(clrScheme.Accent3Color),
                    Accent4: ResolveColor(clrScheme.Accent4Color),
                    Accent5: ResolveColor(clrScheme.Accent5Color),
                    Accent6: ResolveColor(clrScheme.Accent6Color),
                    Hyperlink: ResolveColor(clrScheme.Hyperlink),
                    FollowedHyperlink: ResolveColor(clrScheme.FollowedHyperlinkColor));
            }

            ThemeFontScheme? fontScheme = null;
            var fntScheme = themeElements?.FontScheme;
            if (fntScheme is not null)
            {
                fontScheme = new ThemeFontScheme(
                    Name: fntScheme.Name?.Value,
                    MajorFont: fntScheme.MajorFont?.LatinFont?.Typeface?.Value,
                    MinorFont: fntScheme.MinorFont?.LatinFont?.Typeface?.Value);
            }

            themes.Add(new ThemeInfo(masterIndex, themeName, colorScheme, fontScheme));
            masterIndex++;
        }

        return new ThemeInfoResult(
            Success: true,
            ThemeCount: themes.Count,
            Themes: themes,
            Message: themes.Count == 0
                ? "No slide masters (and therefore no themes) found in the presentation."
                : $"Found {themes.Count} theme(s).");
    }

    /// <summary>
    /// Resolves a DrawingML color element to a #RRGGBB hex string.
    /// Handles <a:srgbClr>, <a:sysClr>, and the container elements used in the color scheme
    /// (e.g. <a:dk1>, <a:dk2>, …).
    /// </summary>
    private static string? ResolveColor(A.Color2Type? colorContainer)
    {
        if (colorContainer is null)
            return null;

        // sRGB hex — most common
        var srgb = colorContainer.RgbColorModelHex;
        if (srgb?.Val?.Value is string srgbVal)
            return "#" + srgbVal.ToUpperInvariant();

        // System color — use the lastClr attribute which stores the resolved value
        var sys = colorContainer.SystemColor;
        if (sys?.LastColor?.Value is string lastClr)
            return "#" + lastClr.ToUpperInvariant();

        return null;
    }
}
