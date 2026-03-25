using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxTools.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Services;

public partial class PresentationService
{
    public List<HyperlinkInfo> GetHyperlinks(string filePath, int? slideNumber = null)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var slideIds = GetSlideIds(doc);
        if (slideIds.Count == 0) return [];

        var results = new List<HyperlinkInfo>();

        for (int i = 0; i < slideIds.Count; i++)
        {
            int currentSlideNumber = i + 1;
            if (slideNumber.HasValue && currentSlideNumber != slideNumber.Value)
                continue;

            var slidePart = GetSlidePart(doc, slideIds, i);
            CollectHyperlinksFromSlide(slidePart, currentSlideNumber, slideIds, doc, results);
        }

        return results;
    }

    public HyperlinkResult AddHyperlink(string filePath, int slideNumber, string shapeName, string url, string? tooltip = null)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var slideIds = GetSlideIds(doc);

        if (slideNumber <= 0 || slideNumber > slideIds.Count)
            throw new ArgumentOutOfRangeException(nameof(slideNumber),
                $"slideNumber {slideNumber} is out of range. Presentation has {slideIds.Count} slide(s).");

        var slidePart = GetSlidePart(doc, slideIds, slideNumber - 1);
        var shape = FindHyperlinkShapeByName(slidePart, shapeName)
            ?? throw new ArgumentException($"Shape '{shapeName}' not found on slide {slideNumber}.");

        var nvProps = GetHyperlinkNvDrawingProps(shape);
        if (nvProps is null)
            throw new InvalidOperationException($"Shape '{shapeName}' does not support hyperlinks.");

        // Create the hyperlink relationship and set HyperlinkOnClick
        var relationship = slidePart.AddHyperlinkRelationship(new Uri(url, UriKind.Absolute), true);
        var hlinkClick = new A.HyperlinkOnClick { Id = relationship.Id };
        if (tooltip is not null)
            hlinkClick.Tooltip = tooltip;

        // Replace any existing hyperlink
        var existing = nvProps.GetFirstChild<A.HyperlinkOnClick>();
        if (existing is not null)
        {
            RemoveHyperlinkRelationship(slidePart, existing.Id?.Value);
            nvProps.ReplaceChild(hlinkClick, existing);
        }
        else
        {
            nvProps.AppendChild(hlinkClick);
        }

        slidePart.Slide.Save();

        return new HyperlinkResult(
            Success: true,
            Action: "Add",
            SlideNumber: slideNumber,
            ShapeName: shapeName,
            Url: url,
            HyperlinkCount: 1,
            Hyperlinks: null,
            Message: $"Hyperlink added to shape '{shapeName}' on slide {slideNumber}.");
    }

    public HyperlinkResult UpdateHyperlink(string filePath, int slideNumber, string shapeName, string newUrl, string? newTooltip = null)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var slideIds = GetSlideIds(doc);

        if (slideNumber <= 0 || slideNumber > slideIds.Count)
            throw new ArgumentOutOfRangeException(nameof(slideNumber),
                $"slideNumber {slideNumber} is out of range. Presentation has {slideIds.Count} slide(s).");

        var slidePart = GetSlidePart(doc, slideIds, slideNumber - 1);
        var shape = FindHyperlinkShapeByName(slidePart, shapeName)
            ?? throw new ArgumentException($"Shape '{shapeName}' not found on slide {slideNumber}.");

        // Try shape-level hyperlink first
        var nvProps = GetHyperlinkNvDrawingProps(shape);
        var shapeHlink = nvProps?.GetFirstChild<A.HyperlinkOnClick>();
        if (shapeHlink is not null)
        {
            UpdateHyperlinkRelationship(slidePart, shapeHlink, newUrl, newTooltip);
            slidePart.Slide.Save();
            return new HyperlinkResult(
                Success: true, Action: "Update", SlideNumber: slideNumber,
                ShapeName: shapeName, Url: newUrl, HyperlinkCount: 1, Hyperlinks: null,
                Message: $"Shape-level hyperlink updated on '{shapeName}' slide {slideNumber}.");
        }

        // Fall back to first run-level hyperlink
        if (shape is Shape textShape)
        {
            var run = textShape.TextBody?.Descendants<A.Run>()
                .FirstOrDefault(r => r.RunProperties?.GetFirstChild<A.HyperlinkOnClick>() is not null);
            var runHlink = run?.RunProperties?.GetFirstChild<A.HyperlinkOnClick>();
            if (runHlink is not null)
            {
                UpdateHyperlinkRelationship(slidePart, runHlink, newUrl, newTooltip);
                slidePart.Slide.Save();
                return new HyperlinkResult(
                    Success: true, Action: "Update", SlideNumber: slideNumber,
                    ShapeName: shapeName, Url: newUrl, HyperlinkCount: 1, Hyperlinks: null,
                    Message: $"Run-level hyperlink updated on '{shapeName}' slide {slideNumber}.");
            }
        }

        throw new InvalidOperationException($"No hyperlink found on shape '{shapeName}' on slide {slideNumber}.");
    }

    public HyperlinkResult RemoveHyperlink(string filePath, int slideNumber, string shapeName)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var slideIds = GetSlideIds(doc);

        if (slideNumber <= 0 || slideNumber > slideIds.Count)
            throw new ArgumentOutOfRangeException(nameof(slideNumber),
                $"slideNumber {slideNumber} is out of range. Presentation has {slideIds.Count} slide(s).");

        var slidePart = GetSlidePart(doc, slideIds, slideNumber - 1);
        var shape = FindHyperlinkShapeByName(slidePart, shapeName)
            ?? throw new ArgumentException($"Shape '{shapeName}' not found on slide {slideNumber}.");

        int removed = 0;

        // Remove shape-level hyperlink
        var nvProps = GetHyperlinkNvDrawingProps(shape);
        var shapeHlink = nvProps?.GetFirstChild<A.HyperlinkOnClick>();
        if (shapeHlink is not null)
        {
            RemoveHyperlinkRelationship(slidePart, shapeHlink.Id?.Value);
            shapeHlink.Remove();
            removed++;
        }

        // Remove run-level hyperlinks
        if (shape is Shape textShape && textShape.TextBody is not null)
        {
            foreach (var run in textShape.TextBody.Descendants<A.Run>().ToList())
            {
                var runHlink = run.RunProperties?.GetFirstChild<A.HyperlinkOnClick>();
                if (runHlink is not null)
                {
                    RemoveHyperlinkRelationship(slidePart, runHlink.Id?.Value);
                    runHlink.Remove();
                    removed++;
                }
            }
        }

        if (removed == 0)
            throw new InvalidOperationException($"No hyperlink found on shape '{shapeName}' on slide {slideNumber}.");

        slidePart.Slide.Save();

        return new HyperlinkResult(
            Success: true,
            Action: "Remove",
            SlideNumber: slideNumber,
            ShapeName: shapeName,
            Url: null,
            HyperlinkCount: removed,
            Hyperlinks: null,
            Message: $"Removed {removed} hyperlink(s) from shape '{shapeName}' on slide {slideNumber}.");
    }

    // --- Hyperlink helpers ---

    private static void CollectHyperlinksFromSlide(
        SlidePart slidePart, int slideNumber,
        IReadOnlyList<SlideId> slideIds, PresentationDocument doc,
        List<HyperlinkInfo> results)
    {
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree is null) return;

        foreach (var shape in shapeTree.ChildElements)
        {
            var nvProps = GetHyperlinkNvDrawingProps(shape);
            if (nvProps is null) continue;
            var shapeName = nvProps.Name?.Value ?? "";

            // Shape-level hyperlink
            var shapeHlink = nvProps.GetFirstChild<A.HyperlinkOnClick>();
            if (shapeHlink is not null)
            {
                var info = ResolveHyperlinkInfo(
                    slidePart, slideIds, doc, slideNumber, shapeName,
                    shapeHlink, displayText: null);
                if (info is not null) results.Add(info);
            }

            // Run-level hyperlinks (only on shapes with text bodies)
            if (shape is Shape textShape && textShape.TextBody is not null)
            {
                foreach (var run in textShape.TextBody.Descendants<A.Run>())
                {
                    var runHlink = run.RunProperties?.GetFirstChild<A.HyperlinkOnClick>();
                    if (runHlink is not null)
                    {
                        var info = ResolveHyperlinkInfo(
                            slidePart, slideIds, doc, slideNumber, shapeName,
                            runHlink, displayText: run.Text?.Text);
                        if (info is not null) results.Add(info);
                    }
                }
            }
        }
    }

    private static HyperlinkInfo? ResolveHyperlinkInfo(
        SlidePart slidePart,
        IReadOnlyList<SlideId> slideIds,
        PresentationDocument doc,
        int slideNumber,
        string shapeName,
        A.HyperlinkOnClick hlinkClick,
        string? displayText)
    {
        var relId = hlinkClick.Id?.Value;
        var tooltip = hlinkClick.Tooltip?.Value;
        var action = hlinkClick.Action?.Value;

        // Internal slide link
        if (action is not null && action.Contains("hlinksldjump", StringComparison.OrdinalIgnoreCase))
        {
            int? targetSlide = ResolveTargetSlideNumber(slidePart, relId, slideIds, doc);
            return new HyperlinkInfo(slideNumber, shapeName, displayText,
                Url: null, TargetSlideNumber: targetSlide, Tooltip: tooltip, HyperlinkType: "internal");
        }

        // External or email link
        if (!string.IsNullOrEmpty(relId))
        {
            try
            {
                var rel = slidePart.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
                if (rel is not null)
                {
                    var url = rel.Uri.OriginalString;
                    var type = url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) ? "email" : "external";
                    return new HyperlinkInfo(slideNumber, shapeName, displayText,
                        Url: url, TargetSlideNumber: null, Tooltip: tooltip, HyperlinkType: type);
                }
            }
            catch
            {
                // Relationship not found — skip gracefully
            }
        }

        return null;
    }

    private static int? ResolveTargetSlideNumber(
        SlidePart slidePart, string? relId,
        IReadOnlyList<SlideId> slideIds, PresentationDocument doc)
    {
        if (string.IsNullOrEmpty(relId)) return null;

        try
        {
            var targetPart = slidePart.GetPartById(relId);
            if (targetPart is SlidePart targetSlidePart)
            {
                var targetUri = targetSlidePart.Uri.ToString();
                for (int i = 0; i < slideIds.Count; i++)
                {
                    var candidatePart = (SlidePart)doc.PresentationPart!.GetPartById(slideIds[i].RelationshipId!.Value!);
                    if (candidatePart.Uri.ToString() == targetUri)
                        return i + 1;
                }
            }
        }
        catch
        {
            // Broken relationship — return null
        }

        return null;
    }

    /// <summary>
    /// Get the non-visual drawing properties element from any presentation shape type.
    /// Returns the P-namespace NonVisualDrawingProperties which has Name, Id, and child HyperlinkOnClick elements.
    /// </summary>
    private static P.NonVisualDrawingProperties? GetHyperlinkNvDrawingProps(OpenXmlElement shape)
    {
        return shape switch
        {
            Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties,
            P.Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties,
            P.ConnectionShape c => c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties,
            P.GroupShape g => g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
            P.GraphicFrame gf => gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
            _ => null
        };
    }

    private static OpenXmlElement? FindHyperlinkShapeByName(SlidePart slidePart, string shapeName)
    {
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree is null) return null;

        foreach (var shape in shapeTree.ChildElements)
        {
            var nvProps = GetHyperlinkNvDrawingProps(shape);
            if (nvProps is not null &&
                string.Equals(nvProps.Name?.Value, shapeName, StringComparison.OrdinalIgnoreCase))
            {
                return shape;
            }
        }

        return null;
    }

    private static void UpdateHyperlinkRelationship(SlidePart slidePart, A.HyperlinkOnClick hlinkClick, string newUrl, string? newTooltip)
    {
        var oldRelId = hlinkClick.Id?.Value;
        RemoveHyperlinkRelationship(slidePart, oldRelId);

        var newRel = slidePart.AddHyperlinkRelationship(new Uri(newUrl, UriKind.Absolute), true);
        hlinkClick.Id = newRel.Id;

        // Clear any existing Action attribute (e.g. "ppaction://hlinksldjump" for internal slide
        // jumps) so the updated link is treated as a plain external hyperlink by GetHyperlinks
        // and by PowerPoint itself.
        hlinkClick.Action = null;

        if (newTooltip is not null)
            hlinkClick.Tooltip = newTooltip;
    }

    private static void RemoveHyperlinkRelationship(SlidePart slidePart, string? relId)
    {
        if (string.IsNullOrEmpty(relId)) return;
        try
        {
            slidePart.DeleteReferenceRelationship(relId);
        }
        catch
        {
            // Relationship already gone — nothing to clean up
        }
    }
}
