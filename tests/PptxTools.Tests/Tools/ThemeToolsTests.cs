using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxTools.Tests.Tools;

/// <summary>
/// Tool-level tests for pptx_inspect_theme.
/// Validates JSON structure, color/font extraction, error handling, and multi-master presentations.
/// </summary>
[Trait("Category", "Integration")]
public class ThemeToolsTests : PptxTestBase
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public ThemeToolsTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    // ────────────────────────────────────────────────────────
    // File not found: structured error
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task InspectTheme_FileNotFound_ReturnsError()
    {
        var fakePath = @"C:\does-not-exist\missing.pptx";

        var result = await _tools.pptx_inspect_theme(fakePath);

        var parsed = JsonSerializer.Deserialize<ThemeInfoResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("File not found", parsed.Message);
        Assert.Equal(0, parsed.ThemeCount);
        Assert.Empty(parsed.Themes);
    }

    // ────────────────────────────────────────────────────────
    // No theme part: still returns success with zero themes
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task InspectTheme_NoThemePart_ReturnsSuccessWithNoColors()
    {
        // Minimal PPTX has a slide master but no ThemePart attached to it.
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_inspect_theme(path);

        var parsed = JsonSerializer.Deserialize<ThemeInfoResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Single(parsed.Themes);
        Assert.Null(parsed.Themes[0].ThemeName);
        Assert.Null(parsed.Themes[0].ColorScheme);
        Assert.Null(parsed.Themes[0].FontScheme);
    }

    // ────────────────────────────────────────────────────────
    // JSON shape: all expected fields present
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task InspectTheme_ResponseJson_HasTopLevelFields()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_inspect_theme(path);

        using var doc = JsonDocument.Parse(result);
        var root = doc.RootElement;
        Assert.True(root.TryGetProperty("Success", out _));
        Assert.True(root.TryGetProperty("ThemeCount", out _));
        Assert.True(root.TryGetProperty("Themes", out _));
        Assert.True(root.TryGetProperty("Message", out _));
    }

    [Fact]
    public async Task InspectTheme_ResponseJson_IsIndented()
    {
        var path = CreateMinimalPptx();

        var result = await _tools.pptx_inspect_theme(path);

        Assert.Contains(Environment.NewLine, result);
    }

    // ────────────────────────────────────────────────────────
    // With a real theme part: colors and fonts extracted
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task InspectTheme_WithThemePart_ReturnsColorScheme()
    {
        var path = CreatePptxWithTheme();

        var result = await _tools.pptx_inspect_theme(path);

        var parsed = JsonSerializer.Deserialize<ThemeInfoResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Single(parsed.Themes);

        var theme = parsed.Themes[0];
        Assert.Equal("TestTheme", theme.ThemeName);

        var cs = theme.ColorScheme;
        Assert.NotNull(cs);
        Assert.Equal("TestColors", cs.Name);
        Assert.Equal("#000000", cs.Dark1);
        Assert.Equal("#FFFFFF", cs.Light1);
        Assert.Equal("#1F1E32", cs.Dark2);
        Assert.Equal("#A5A6A6", cs.Light2);
        Assert.Equal("#FF0000", cs.Accent1);
        Assert.Equal("#00FF00", cs.Accent2);
        Assert.Equal("#0000FF", cs.Accent3);
        Assert.Equal("#FFFF00", cs.Accent4);
        Assert.Equal("#FF00FF", cs.Accent5);
        Assert.Equal("#00FFFF", cs.Accent6);
        Assert.Equal("#800080", cs.Hyperlink);
        Assert.Equal("#C0C0C0", cs.FollowedHyperlink);
    }

    [Fact]
    public async Task InspectTheme_WithThemePart_ReturnsFontScheme()
    {
        var path = CreatePptxWithTheme();

        var result = await _tools.pptx_inspect_theme(path);

        var parsed = JsonSerializer.Deserialize<ThemeInfoResult>(result, JsonOptions);
        Assert.NotNull(parsed);

        var fs = parsed.Themes[0].FontScheme;
        Assert.NotNull(fs);
        Assert.Equal("TestFonts", fs.Name);
        Assert.Equal("Calibri Light", fs.MajorFont);
        Assert.Equal("Calibri", fs.MinorFont);
    }

    [Fact]
    public async Task InspectTheme_WithThemePart_MasterIndexIsZero()
    {
        var path = CreatePptxWithTheme();

        var result = await _tools.pptx_inspect_theme(path);

        var parsed = JsonSerializer.Deserialize<ThemeInfoResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.Equal(0, parsed.Themes[0].MasterIndex);
    }

    // ────────────────────────────────────────────────────────
    // System colors (sysClr lastClr)
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task InspectTheme_SystemColors_ResolvedFromLastClr()
    {
        var path = CreatePptxWithSystemColorTheme();

        var result = await _tools.pptx_inspect_theme(path);

        var parsed = JsonSerializer.Deserialize<ThemeInfoResult>(result, JsonOptions);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);

        var cs = parsed.Themes[0].ColorScheme;
        Assert.NotNull(cs);
        // Dark1 comes from sysClr lastClr="000000"
        Assert.Equal("#000000", cs.Dark1);
        // Light1 from sysClr lastClr="FFFFFF"
        Assert.Equal("#FFFFFF", cs.Light1);
    }

    // ────────────────────────────────────────────────────────
    // Fixture helpers
    // ────────────────────────────────────────────────────────

    private string CreatePptxWithTheme()
    {
        var path = CreateMinimalPptx();

        using var doc = PresentationDocument.Open(path, true);
        var masterPart = doc.PresentationPart!.SlideMasterParts.First();
        var themePart = masterPart.AddNewPart<ThemePart>();

        themePart.Theme = BuildTheme(
            themeName: "TestTheme",
            colorSchemeName: "TestColors",
            dark1: "000000",
            light1: "FFFFFF",
            dark2: "1F1E32",
            light2: "A5A6A6",
            accent1: "FF0000",
            accent2: "00FF00",
            accent3: "0000FF",
            accent4: "FFFF00",
            accent5: "FF00FF",
            accent6: "00FFFF",
            hyperlink: "800080",
            followedHyperlink: "C0C0C0",
            fontSchemeName: "TestFonts",
            majorFont: "Calibri Light",
            minorFont: "Calibri");

        themePart.Theme.Save();
        return path;
    }

    private string CreatePptxWithSystemColorTheme()
    {
        var path = CreateMinimalPptx();

        using var doc = PresentationDocument.Open(path, true);
        var masterPart = doc.PresentationPart!.SlideMasterParts.First();
        var themePart = masterPart.AddNewPart<ThemePart>();

        var colorScheme = new A.ColorScheme { Name = "SysColors" };
        colorScheme.Append(new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }));
        colorScheme.Append(new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }));
        colorScheme.Append(new A.Dark2Color(new A.RgbColorModelHex { Val = "1F1E32" }));
        colorScheme.Append(new A.Light2Color(new A.RgbColorModelHex { Val = "A5A6A6" }));
        colorScheme.Append(new A.Accent1Color(new A.RgbColorModelHex { Val = "FF0000" }));
        colorScheme.Append(new A.Accent2Color(new A.RgbColorModelHex { Val = "00FF00" }));
        colorScheme.Append(new A.Accent3Color(new A.RgbColorModelHex { Val = "0000FF" }));
        colorScheme.Append(new A.Accent4Color(new A.RgbColorModelHex { Val = "FFFF00" }));
        colorScheme.Append(new A.Accent5Color(new A.RgbColorModelHex { Val = "FF00FF" }));
        colorScheme.Append(new A.Accent6Color(new A.RgbColorModelHex { Val = "00FFFF" }));
        colorScheme.Append(new A.Hyperlink(new A.RgbColorModelHex { Val = "800080" }));
        colorScheme.Append(new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "C0C0C0" }));

        var fontScheme = new A.FontScheme { Name = "Office" };
        fontScheme.Append(new A.MajorFont(new A.LatinFont { Typeface = "Calibri Light" }));
        fontScheme.Append(new A.MinorFont(new A.LatinFont { Typeface = "Calibri" }));

        var formatScheme = new A.FormatScheme { Name = "Office" };
        formatScheme.Append(new A.FillStyleList(
            new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })));
        formatScheme.Append(new A.LineStyleList(
            new A.Outline { Width = 6350 }));
        formatScheme.Append(new A.EffectStyleList(
            new A.EffectStyle(new A.EffectList())));
        formatScheme.Append(new A.BackgroundFillStyleList(
            new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })));

        themePart.Theme = new A.Theme { Name = "SysColorTheme" };
        themePart.Theme.Append(new A.ThemeElements(colorScheme, fontScheme, formatScheme));
        themePart.Theme.Save();
        return path;
    }

    private static A.Theme BuildTheme(
        string themeName,
        string colorSchemeName,
        string dark1, string light1, string dark2, string light2,
        string accent1, string accent2, string accent3,
        string accent4, string accent5, string accent6,
        string hyperlink, string followedHyperlink,
        string fontSchemeName, string majorFont, string minorFont)
    {
        var colorScheme = new A.ColorScheme { Name = colorSchemeName };
        colorScheme.Append(new A.Dark1Color(new A.RgbColorModelHex { Val = dark1 }));
        colorScheme.Append(new A.Light1Color(new A.RgbColorModelHex { Val = light1 }));
        colorScheme.Append(new A.Dark2Color(new A.RgbColorModelHex { Val = dark2 }));
        colorScheme.Append(new A.Light2Color(new A.RgbColorModelHex { Val = light2 }));
        colorScheme.Append(new A.Accent1Color(new A.RgbColorModelHex { Val = accent1 }));
        colorScheme.Append(new A.Accent2Color(new A.RgbColorModelHex { Val = accent2 }));
        colorScheme.Append(new A.Accent3Color(new A.RgbColorModelHex { Val = accent3 }));
        colorScheme.Append(new A.Accent4Color(new A.RgbColorModelHex { Val = accent4 }));
        colorScheme.Append(new A.Accent5Color(new A.RgbColorModelHex { Val = accent5 }));
        colorScheme.Append(new A.Accent6Color(new A.RgbColorModelHex { Val = accent6 }));
        colorScheme.Append(new A.Hyperlink(new A.RgbColorModelHex { Val = hyperlink }));
        colorScheme.Append(new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = followedHyperlink }));

        var fontScheme = new A.FontScheme { Name = fontSchemeName };
        fontScheme.Append(new A.MajorFont(new A.LatinFont { Typeface = majorFont }));
        fontScheme.Append(new A.MinorFont(new A.LatinFont { Typeface = minorFont }));

        // FormatScheme is required for a valid theme but we keep it minimal
        var formatScheme = new A.FormatScheme { Name = "Office" };
        formatScheme.Append(new A.FillStyleList(
            new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })));
        formatScheme.Append(new A.LineStyleList(
            new A.Outline { Width = 6350 }));
        formatScheme.Append(new A.EffectStyleList(
            new A.EffectStyle(new A.EffectList())));
        formatScheme.Append(new A.BackgroundFillStyleList(
            new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })));

        var theme = new A.Theme { Name = themeName };
        theme.Append(new A.ThemeElements(colorScheme, fontScheme, formatScheme));
        return theme;
    }
}
