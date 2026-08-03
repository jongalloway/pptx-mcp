using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxTools.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxTools.Services;

public partial class PresentationService
{
    /// <summary>Fill element local names in the DrawingML namespace.</summary>
    private static readonly HashSet<string> FillLocalNames =
        ["solidFill", "noFill", "gradFill", "pattFill", "blipFill", "grpFill"];

    public TextFormattingResult GetTextFormatting(string filePath, int? slideNumber = null, string? shapeName = null)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var slideIds = GetSlideIds(doc);

        if (slideIds.Count == 0)
            return new TextFormattingResult(true, "Get", slideNumber, shapeName, 0, [], "Presentation has no slides.");

        if (slideNumber.HasValue && (slideNumber.Value < 1 || slideNumber.Value > slideIds.Count))
            return new TextFormattingResult(false, "Get", slideNumber, shapeName, 0, [],
                $"slideNumber {slideNumber} is out of range. Presentation has {slideIds.Count} slide(s).");

        var formattings = new List<TextFormattingInfo>();

        int startSlide = slideNumber.HasValue ? slideNumber.Value : 1;
        int endSlide = slideNumber.HasValue ? slideNumber.Value : slideIds.Count;

        for (int sn = startSlide; sn <= endSlide; sn++)
        {
            var slidePart = GetSlidePart(doc, slideIds, sn - 1);
            var slide = slidePart.Slide;
            if (slide?.CommonSlideData?.ShapeTree is null)
                continue;

            foreach (var shape in slide.CommonSlideData.ShapeTree.Elements<Shape>())
            {
                var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "";

                if (!string.IsNullOrWhiteSpace(shapeName) &&
                    !string.Equals(name, shapeName.Trim(), StringComparison.OrdinalIgnoreCase))
                    continue;

                var textBody = shape.TextBody;
                if (textBody is null)
                    continue;

                foreach (var paragraph in textBody.Elements<A.Paragraph>())
                {
                    var alignVal = paragraph.ParagraphProperties?.Alignment?.Value;
                    string? alignment = null;
                    if (alignVal is not null)
                    {
                        if (alignVal == A.TextAlignmentTypeValues.Left) alignment = "Left";
                        else if (alignVal == A.TextAlignmentTypeValues.Center) alignment = "Center";
                        else if (alignVal == A.TextAlignmentTypeValues.Right) alignment = "Right";
                        else if (alignVal == A.TextAlignmentTypeValues.Justified) alignment = "Justify";
                    }

                    foreach (var run in paragraph.Elements<A.Run>())
                    {
                        var rp = run.RunProperties;
                        var text = run.Text?.Text;

                        string? fontFamily = rp?.GetFirstChild<A.LatinFont>()?.Typeface?.Value;
                        double? fontSize = rp?.FontSize is not null ? rp.FontSize.Value / 100.0 : null;
                        bool? bold = rp?.Bold?.Value;
                        bool? italic = rp?.Italic?.Value;
                        bool? underline = rp?.Underline?.HasValue == true
                            ? rp.Underline.Value != A.TextUnderlineValues.None
                            : null;
                        string? color = GetRunColor(rp);

                        formattings.Add(new TextFormattingInfo(
                            SlideNumber: sn,
                            ShapeName: name,
                            Text: text,
                            FontFamily: fontFamily,
                            FontSize: fontSize,
                            Bold: bold,
                            Italic: italic,
                            Underline: underline,
                            Color: color,
                            Alignment: alignment));
                    }
                }
            }
        }

        return new TextFormattingResult(true, "Get", slideNumber, shapeName, formattings.Count, formattings,
            $"Found {formattings.Count} formatted text run(s).");
    }

    public TextFormattingResult ApplyTextFormatting(
        string filePath,
        int slideNumber,
        string shapeName,
        string? fontFamily = null,
        double? fontSize = null,
        bool? bold = null,
        bool? italic = null,
        bool? underline = null,
        string? color = null,
        string? alignment = null)
    {
        // Require at least one formatting property to avoid silent no-ops
        if (!HasAnyFormattingProperties(fontFamily, fontSize, bold, italic, underline, color, alignment))
            return new TextFormattingResult(false, "Apply", slideNumber, shapeName, 0, [],
                "No formatting properties specified. Provide at least one of: fontFamily, fontSize, bold, italic, underline, color, alignment.");

        if (fontSize.HasValue && (fontSize.Value <= 0 || fontSize.Value > int.MaxValue / 100.0))
            return new TextFormattingResult(false, "Apply", slideNumber, shapeName, 0, [],
                $"fontSize must be a positive number no greater than {(int)(int.MaxValue / 100.0)} points.");

        using var doc = PresentationDocument.Open(filePath, true);
        var slideIds = GetSlideIds(doc);

        if (slideNumber < 1 || slideNumber > slideIds.Count)
            return new TextFormattingResult(false, "Apply", slideNumber, shapeName, 0, [],
                $"slideNumber {slideNumber} is out of range. Presentation has {slideIds.Count} slide(s).");

        var slidePart = GetSlidePart(doc, slideIds, slideNumber - 1);
        var slide = slidePart.Slide;
        if (slide?.CommonSlideData?.ShapeTree is null)
            return new TextFormattingResult(false, "Apply", slideNumber, shapeName, 0, [],
                $"Slide {slideNumber} has no shape tree.");

        var targetShape = slide.CommonSlideData.ShapeTree.Elements<Shape>()
            .FirstOrDefault(s =>
                string.Equals(
                    s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                    shapeName.Trim(),
                    StringComparison.OrdinalIgnoreCase));

        if (targetShape is null)
        {
            var available = string.Join(", ",
                slide.CommonSlideData.ShapeTree.Elements<Shape>()
                    .Select(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value)
                    .Where(n => !string.IsNullOrWhiteSpace(n)));
            return new TextFormattingResult(false, "Apply", slideNumber, shapeName, 0, [],
                $"No shape named '{shapeName}' found on slide {slideNumber}. Available shapes: {available}");
        }

        var textBody = targetShape.TextBody;
        if (textBody is null)
            return new TextFormattingResult(false, "Apply", slideNumber, shapeName, 0, [],
                $"Shape '{shapeName}' on slide {slideNumber} has no text body.");

        // Apply alignment at paragraph level
        if (!string.IsNullOrWhiteSpace(alignment))
        {
            var alignValue = ParseAlignment(alignment);
            foreach (var paragraph in textBody.Elements<A.Paragraph>())
            {
                var pp = paragraph.ParagraphProperties;
                if (pp is null)
                {
                    pp = new A.ParagraphProperties();
                    paragraph.InsertAt(pp, 0);
                }
                pp.Alignment = alignValue;
            }
        }

        // Apply run-level formatting
        int runsModified = 0;
        foreach (var paragraph in textBody.Elements<A.Paragraph>())
        {
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var rp = run.RunProperties;
                if (rp is null)
                {
                    rp = new A.RunProperties();
                    run.InsertAt(rp, 0);
                }

                var runModified = false;

                if (bold.HasValue && (rp.Bold is null || rp.Bold.Value != bold.Value))
                {
                    rp.Bold = bold.Value;
                    runModified = true;
                }

                if (italic.HasValue && (rp.Italic is null || rp.Italic.Value != italic.Value))
                {
                    rp.Italic = italic.Value;
                    runModified = true;
                }

                if (underline.HasValue)
                {
                    var desiredUnderline = underline.Value
                        ? A.TextUnderlineValues.Single
                        : A.TextUnderlineValues.None;
                    if (rp.Underline is null || rp.Underline.Value != desiredUnderline)
                    {
                        rp.Underline = desiredUnderline;
                        runModified = true;
                    }
                }

                if (fontSize.HasValue)
                {
                    var newFontSize = (int)Math.Round(fontSize.Value * 100, MidpointRounding.AwayFromZero);
                    if (rp.FontSize is null || rp.FontSize.Value != newFontSize)
                    {
                        rp.FontSize = newFontSize;
                        runModified = true;
                    }
                }

                if (!string.IsNullOrWhiteSpace(fontFamily))
                {
                    var trimmedFontFamily = fontFamily.Trim();
                    var existingLatin = rp.GetFirstChild<A.LatinFont>();
                    if (existingLatin?.Typeface?.Value != trimmedFontFamily)
                    {
                        if (existingLatin is not null)
                            rp.RemoveChild(existingLatin);
                        rp.Append(new A.LatinFont { Typeface = trimmedFontFamily });
                        runModified = true;
                    }
                }

                if (!string.IsNullOrWhiteSpace(color))
                {
                    var normalizedHex = NormalizeHexColor(color.Trim()); // validates; throws if invalid
                    var existingColor = GetRunColor(rp);
                    var existingHex = existingColor?.StartsWith('#') == true ? existingColor[1..] : null;
                    if (!string.Equals(existingHex, normalizedHex, StringComparison.OrdinalIgnoreCase))
                    {
                        ApplyRunColor(rp, normalizedHex);
                        runModified = true;
                    }
                }

                if (runModified)
                    runsModified++;
            }
        }

        slide.Save();

        return new TextFormattingResult(true, "Apply", slideNumber, shapeName, runsModified, [],
            $"Applied formatting to {runsModified} text run(s) in shape '{shapeName}' on slide {slideNumber}.");
    }

    private static string? GetRunColor(A.RunProperties? rp)
    {
        if (rp is null) return null;
        var solidFill = rp.GetFirstChild<A.SolidFill>();
        if (solidFill is null) return null;

        var rgbColor = solidFill.GetFirstChild<A.RgbColorModelHex>();
        if (rgbColor?.Val?.Value is not null)
            return $"#{rgbColor.Val.Value}";

        var schemeColor = solidFill.GetFirstChild<A.SchemeColor>();
        if (schemeColor?.Val?.Value is not null)
            return $"scheme:{schemeColor.Val.Value}";

        var systemColor = solidFill.GetFirstChild<A.SystemColor>();
        if (systemColor?.Val?.Value is not null)
            return $"system:{systemColor.Val.Value}";

        var presetColor = solidFill.GetFirstChild<A.PresetColor>();
        if (presetColor?.Val?.Value is not null)
            return $"preset:{presetColor.Val.Value}";

        return null;
    }

    /// <summary>
    /// Validates <paramref name="hexColor"/> and returns the 6-digit uppercase hex without '#'.
    /// Throws <see cref="ArgumentException"/> for invalid input.
    /// </summary>
    private static string NormalizeHexColor(string hexColor)
    {
        var hex = hexColor.StartsWith('#') ? hexColor[1..] : hexColor;
        if (!Regex.IsMatch(hex, @"^[0-9A-Fa-f]{6}$"))
            throw new ArgumentException(
                $"Invalid hex color '{hexColor}'. Expected a 6-digit RGB hex color (e.g. \"#FF0000\" or \"FF0000\").");
        return hex.ToUpperInvariant();
    }

    /// <param name="upperHex">6-digit uppercase hex without '#', already validated by <see cref="NormalizeHexColor"/>.</param>
    private static void ApplyRunColor(A.RunProperties rp, string upperHex)
    {
        // Remove all fill-related children to prevent conflicting fill definitions in the XML
        foreach (var fill in rp.ChildElements.Where(e => FillLocalNames.Contains(e.LocalName)).ToList())
            rp.RemoveChild(fill);

        rp.InsertAt(new A.SolidFill(new A.RgbColorModelHex { Val = upperHex }), 0);
    }

    private static bool HasAnyFormattingProperties(
        string? fontFamily, double? fontSize, bool? bold, bool? italic,
        bool? underline, string? color, string? alignment) =>
        fontFamily is not null || fontSize is not null || bold is not null || italic is not null
        || underline is not null || !string.IsNullOrWhiteSpace(color) || !string.IsNullOrWhiteSpace(alignment);

    private static A.TextAlignmentTypeValues ParseAlignment(string alignment) =>
        alignment.Trim().ToLowerInvariant() switch
        {
            "left" => A.TextAlignmentTypeValues.Left,
            "center" => A.TextAlignmentTypeValues.Center,
            "right" => A.TextAlignmentTypeValues.Right,
            "justify" or "justified" => A.TextAlignmentTypeValues.Justified,
            _ => throw new ArgumentException($"Unknown alignment '{alignment}'. Valid values: Left, Center, Right, Justify.")
        };
}
