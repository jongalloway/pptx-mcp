using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxTools.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxTools.Services;

public partial class PresentationService
{
    public TextFormattingResult GetTextFormatting(string filePath, int? slideNumber = null, string? shapeName = null)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var slideIds = GetSlideIds(doc);

        if (slideIds.Count == 0)
            return new TextFormattingResult(true, "Get", slideNumber, shapeName, 0, [], "Presentation has no slides.");

        var formattings = new List<TextFormattingInfo>();

        int startSlide = slideNumber.HasValue ? slideNumber.Value : 1;
        int endSlide = slideNumber.HasValue ? slideNumber.Value : slideIds.Count;

        if (slideNumber.HasValue && (slideNumber.Value < 1 || slideNumber.Value > slideIds.Count))
            return new TextFormattingResult(false, "Get", slideNumber, shapeName, 0, [],
                $"slideNumber {slideNumber} is out of range. Presentation has {slideIds.Count} slide(s).");

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

                if (bold.HasValue)
                    rp.Bold = bold.Value;

                if (italic.HasValue)
                    rp.Italic = italic.Value;

                if (underline.HasValue)
                    rp.Underline = underline.Value ? A.TextUnderlineValues.Single : A.TextUnderlineValues.None;

                if (fontSize.HasValue)
                    rp.FontSize = (int)(fontSize.Value * 100);

                if (!string.IsNullOrWhiteSpace(fontFamily))
                {
                    var existingLatin = rp.GetFirstChild<A.LatinFont>();
                    if (existingLatin is not null)
                        rp.RemoveChild(existingLatin);
                    rp.Append(new A.LatinFont { Typeface = fontFamily.Trim() });
                }

                if (!string.IsNullOrWhiteSpace(color))
                {
                    ApplyRunColor(rp, color.Trim());
                }

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
        var rgbColor = solidFill?.GetFirstChild<A.RgbColorModelHex>();
        return rgbColor?.Val?.Value is not null ? $"#{rgbColor.Val.Value}" : null;
    }

    private static void ApplyRunColor(A.RunProperties rp, string hexColor)
    {
        var hex = hexColor.StartsWith('#') ? hexColor[1..] : hexColor;

        var existingSolidFill = rp.GetFirstChild<A.SolidFill>();
        if (existingSolidFill is not null)
            rp.RemoveChild(existingSolidFill);

        rp.InsertAt(new A.SolidFill(new A.RgbColorModelHex { Val = hex }), 0);
    }

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
