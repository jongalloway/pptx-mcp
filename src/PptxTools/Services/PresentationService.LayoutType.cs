using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxTools.Models;

namespace PptxTools.Services;

public partial class PresentationService
{
    /// <summary>
    /// Valid layout type strings and their corresponding <see cref="SlideLayoutValues"/> values.
    /// These strings match the XML attribute values used in the OOXML spec.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, SlideLayoutValues> LayoutTypeMap =
        new Dictionary<string, SlideLayoutValues>(StringComparer.OrdinalIgnoreCase)
        {

            ["title"]   = SlideLayoutValues.Title,
            ["tx"]      = SlideLayoutValues.Text,
            ["twoColTx"] = SlideLayoutValues.TwoColumnText,
            ["tbl"]     = SlideLayoutValues.Table,
            ["txAndChart"] = SlideLayoutValues.TextAndChart,
            ["chartAndTx"] = SlideLayoutValues.ChartAndText,
            ["dgm"]     = SlideLayoutValues.Diagram,
            ["chart"]   = SlideLayoutValues.Chart,
            ["txAndClipArt"] = SlideLayoutValues.TextAndClipArt,
            ["clipArtAndTx"] = SlideLayoutValues.ClipArtAndText,
            ["titleOnly"] = SlideLayoutValues.TitleOnly,
            ["blank"]   = SlideLayoutValues.Blank,
            ["txAndObj"] = SlideLayoutValues.TextAndObject,
            ["objAndTx"] = SlideLayoutValues.ObjectAndText,
            ["objOnly"] = SlideLayoutValues.ObjectOnly,
            ["obj"]     = SlideLayoutValues.Object,
            ["txAndMedia"] = SlideLayoutValues.TextAndMedia,
            ["mediaAndTx"] = SlideLayoutValues.MidiaAndText,  // "MidiaAndText" is a known typo in the OpenXML SDK for "MediaAndText"; this is intentional
            ["objOverTx"] = SlideLayoutValues.ObjectOverText,
            ["txOverObj"] = SlideLayoutValues.TextOverObject,
            ["txAndTwoObj"] = SlideLayoutValues.TextAndTwoObjects,
            ["twoObjAndTx"] = SlideLayoutValues.TwoObjectsAndText,
            ["twoObjOverTx"] = SlideLayoutValues.TwoObjectsOverText,
            ["fourObj"] = SlideLayoutValues.FourObjects,
            ["vertTx"]  = SlideLayoutValues.VerticalText,
            ["clipArtAndVertTx"] = SlideLayoutValues.ClipArtAndVerticalText,
            ["vertTitleAndTx"] = SlideLayoutValues.VerticalTitleAndText,
            ["vertTitleAndTxOverChart"] = SlideLayoutValues.VerticalTitleAndTextOverChart,
            ["twoObj"] = SlideLayoutValues.TwoObjects,
            ["objAndTwoObj"] = SlideLayoutValues.ObjectAndTwoObjects,
            ["twoObjAndObj"] = SlideLayoutValues.TwoObjectsAndObject,
            ["cust"]    = SlideLayoutValues.Custom,
            ["secHead"] = SlideLayoutValues.SectionHeader,
            ["twoTxTwoObj"] = SlideLayoutValues.TwoTextAndTwoObjects,
            ["objTx"]   = SlideLayoutValues.ObjectText,
            ["picTx"]   = SlideLayoutValues.PictureText,
        };

    /// <summary>
    /// Reverse of <see cref="LayoutTypeMap"/>: maps <see cref="SlideLayoutValues"/> → OOXML attribute string.
    /// </summary>
    private static readonly IReadOnlyDictionary<SlideLayoutValues, string> LayoutTypeReverseMap =
        LayoutTypeMap.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    /// <summary>
    /// Set the semantic <c>type</c> attribute on the named slide layout in a PPTX file.
    /// </summary>
    /// <param name="filePath">Path to the .pptx file to modify.</param>
    /// <param name="layoutName">Display name of the layout to update (case-insensitive).</param>
    /// <param name="layoutType">
    /// The OOXML type string to set (e.g. "title", "obj", "secHead", "blank", "twoObj", "picTx").
    /// See <see cref="LayoutTypeMap"/> for the full list of accepted values.
    /// </param>
    /// <returns>A result record describing the outcome.</returns>
    public SetLayoutTypeResult SetLayoutType(string filePath, string layoutName, string layoutType)
    {
        if (!LayoutTypeMap.TryGetValue(layoutType, out var sdkValue))
        {
            var valid = string.Join(", ", LayoutTypeMap.Keys.Order());
            return new SetLayoutTypeResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                LayoutType: layoutType,
                Message: $"Unknown layout type '{layoutType}'. Valid values: {valid}.");
        }

        using var doc = PresentationDocument.Open(filePath, isEditable: true);
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation part is missing.");

        SlideLayoutPart? targetPart = null;
        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            foreach (var layoutPart in masterPart.SlideLayoutParts)
            {
                var name = layoutPart.SlideLayout?.CommonSlideData?.Name?.Value ?? string.Empty;
                if (string.Equals(name, layoutName, StringComparison.OrdinalIgnoreCase))
                {
                    targetPart = layoutPart;
                    break;
                }
            }
            if (targetPart is not null) break;
        }

        if (targetPart is null)
        {
            return new SetLayoutTypeResult(
                Success: false,
                FilePath: filePath,
                LayoutName: layoutName,
                LayoutType: layoutType,
                Message: $"Layout '{layoutName}' was not found in the presentation.");
        }

        targetPart.SlideLayout.Type = sdkValue;
        targetPart.SlideLayout.Save();

        return new SetLayoutTypeResult(
            Success: true,
            FilePath: filePath,
            LayoutName: layoutName,
            LayoutType: layoutType,
            Message: $"Set type '{layoutType}' on layout '{layoutName}'.");
    }
}
