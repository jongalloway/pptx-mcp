using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxTools.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Services;

public partial class PresentationService
{
    public ValidationResult ValidatePresentation(string filePath, int? slideNumber = null)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var slideIds = GetSlideIds(doc);

        if (slideIds.Count == 0)
        {
            return new ValidationResult(
                Success: true,
                Action: "Validate",
                IssueCount: 0,
                ErrorCount: 0,
                WarningCount: 0,
                InfoCount: 0,
                Issues: [],
                Message: "Presentation has no slides.");
        }

        if (slideNumber.HasValue && (slideNumber.Value < 1 || slideNumber.Value > slideIds.Count))
        {
            return new ValidationResult(
                Success: false,
                Action: "Validate",
                IssueCount: 0,
                ErrorCount: 0,
                WarningCount: 0,
                InfoCount: 0,
                Issues: [],
                Message: $"slideNumber {slideNumber.Value} is out of range. Presentation has {slideIds.Count} slide(s).");
        }

        var issues = new List<ValidationIssue>();

        for (int i = 0; i < slideIds.Count; i++)
        {
            int currentSlideNumber = i + 1;
            if (slideNumber.HasValue && currentSlideNumber != slideNumber.Value)
                continue;

            var slidePart = GetSlidePart(doc, slideIds, i);

            // Check for corrupt XML in relationship parts before touching slide XML
            CheckCorruptPartXml(slidePart, currentSlideNumber, issues);

            try
            {
                // All checks that access slidePart.Slide — wrapping catches corrupt slide XML
                CheckRequiredElements(slidePart, currentSlideNumber, issues);
                CheckDuplicateShapeIds(slidePart, currentSlideNumber, issues);
                CheckMissingImageReferences(slidePart, currentSlideNumber, issues);
                CheckOrphanedRelationships(slidePart, currentSlideNumber, issues);
                CheckHyperlinkTargets(slidePart, currentSlideNumber, issues);
            }
            catch (XmlException ex)
            {
                issues.Add(new ValidationIssue(
                    SlideNumber: currentSlideNumber,
                    Severity: ValidationSeverity.Error,
                    Category: "CorruptSlideXml",
                    Description: $"Slide {currentSlideNumber} XML is malformed and cannot be fully validated.",
                    Recommendation: "This slide may be corrupt. Consider recreating it from a layout.",
                    XmlContext: $"Line {ex.LineNumber}, Position {ex.LinePosition}"));
            }
        }

        // Presentation-wide: check for duplicate shape IDs across slides (only when not filtering)
        if (!slideNumber.HasValue && slideIds.Count > 1)
            CheckCrossSlideShapeIdDuplicates(doc, slideIds, issues);

        // Sort: errors first, then warnings, then info
        issues.Sort((a, b) => a.Severity.CompareTo(b.Severity));

        int errorCount = issues.Count(i => i.Severity == ValidationSeverity.Error);
        int warningCount = issues.Count(i => i.Severity == ValidationSeverity.Warning);
        int infoCount = issues.Count(i => i.Severity == ValidationSeverity.Info);

        var summary = issues.Count == 0
            ? "Validation passed — no issues found."
            : $"Found {issues.Count} issue(s): {errorCount} error(s), {warningCount} warning(s), {infoCount} info.";

        return new ValidationResult(
            Success: true,
            Action: "Validate",
            IssueCount: issues.Count,
            ErrorCount: errorCount,
            WarningCount: warningCount,
            InfoCount: infoCount,
            Issues: issues,
            Message: summary);
    }

    private static void CheckRequiredElements(SlidePart slidePart, int slideNumber, List<ValidationIssue> issues)
    {
        if (slidePart.Slide.CommonSlideData is null)
        {
            issues.Add(new ValidationIssue(
                SlideNumber: slideNumber,
                Severity: ValidationSeverity.Error,
                Category: "MissingRequiredElement",
                Description: $"Slide {slideNumber} is missing CommonSlideData (p:cSld).",
                Recommendation: "This slide may be corrupt. Consider recreating it from a layout.",
                XmlContext: "p:sld/p:cSld"));
            return;
        }

        if (slidePart.Slide.CommonSlideData.ShapeTree is null)
        {
            issues.Add(new ValidationIssue(
                SlideNumber: slideNumber,
                Severity: ValidationSeverity.Error,
                Category: "MissingRequiredElement",
                Description: $"Slide {slideNumber} is missing ShapeTree (p:spTree) inside CommonSlideData.",
                Recommendation: "The slide has no shape container. Consider recreating it from a layout.",
                XmlContext: "p:sld/p:cSld/p:spTree"));
        }
    }

    private static void CheckDuplicateShapeIds(SlidePart slidePart, int slideNumber, List<ValidationIssue> issues)
    {
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree is null) return;

        var seenIds = new Dictionary<uint, string>();

        foreach (var child in shapeTree.ChildElements)
        {
            var (id, name) = GetShapeIdAndName(child);
            if (id is null) continue;

            if (seenIds.TryGetValue(id.Value, out var existingName))
            {
                issues.Add(new ValidationIssue(
                    SlideNumber: slideNumber,
                    Severity: ValidationSeverity.Error,
                    Category: "DuplicateShapeId",
                    Description: $"Slide {slideNumber}: Shape ID {id.Value} is shared by '{existingName}' and '{name}'.",
                    Recommendation: "Assign unique IDs to each shape. Duplicate IDs can cause editing failures in PowerPoint.",
                    XmlContext: $"p:sld/p:cSld/p:spTree/{child.LocalName}[cNvPr[@id='{id}']]"));
            }
            else
            {
                seenIds[id.Value] = name;
            }
        }
    }

    private static void CheckMissingImageReferences(SlidePart slidePart, int slideNumber, List<ValidationIssue> issues)
    {
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree is null) return;

        foreach (var picture in shapeTree.Descendants<P.Picture>())
        {
            var blipFill = picture.BlipFill;
            var blip = blipFill?.Blip;
            var embed = blip?.Embed?.Value;

            if (string.IsNullOrEmpty(embed)) continue;

            try
            {
                slidePart.GetPartById(embed);
            }
            catch
            {
                AddMissingImageIssue(slideNumber, embed, picture, issues);
            }
        }
    }

    private static void AddMissingImageIssue(int slideNumber, string relId, P.Picture picture, List<ValidationIssue> issues)
    {
        var shapeName = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "(unnamed)";
        issues.Add(new ValidationIssue(
            SlideNumber: slideNumber,
            Severity: ValidationSeverity.Error,
            Category: "MissingImageReference",
            Description: $"Slide {slideNumber}: Picture '{shapeName}' references relationship '{relId}' which does not resolve to an image part.",
            Recommendation: "Replace the image or remove the picture shape. The image will appear broken in PowerPoint.",
            XmlContext: $"p:sld/p:cSld/p:spTree/p:pic/p:blipFill/a:blip[@r:embed='{relId}']"));
    }

    private static void CheckOrphanedRelationships(SlidePart slidePart, int slideNumber, List<ValidationIssue> issues)
    {
        // Collect all relationship IDs actually referenced in the slide XML
        var referencedRelIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Blip embeds (images)
        foreach (var blip in slidePart.Slide.Descendants<A.Blip>())
        {
            if (!string.IsNullOrEmpty(blip.Embed?.Value))
                referencedRelIds.Add(blip.Embed.Value);
            if (!string.IsNullOrEmpty(blip.Link?.Value))
                referencedRelIds.Add(blip.Link.Value);
        }

        // Hyperlink references
        foreach (var hlink in slidePart.Slide.Descendants<A.HyperlinkOnClick>())
        {
            if (!string.IsNullOrEmpty(hlink.Id?.Value))
                referencedRelIds.Add(hlink.Id.Value);
        }

        foreach (var hlink in slidePart.Slide.Descendants<A.HyperlinkOnMouseOver>())
        {
            if (!string.IsNullOrEmpty(hlink.Id?.Value))
                referencedRelIds.Add(hlink.Id.Value);
        }

        // Video/audio references
        foreach (var videoFile in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.VideoFromFile>())
        {
            if (!string.IsNullOrEmpty(videoFile.Link?.Value))
                referencedRelIds.Add(videoFile.Link.Value);
        }

        foreach (var audioFile in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.AudioFromFile>())
        {
            if (!string.IsNullOrEmpty(audioFile.Link?.Value))
                referencedRelIds.Add(audioFile.Link.Value);
        }

        // Check for parts that have no reference in the slide XML
        foreach (var rel in slidePart.Parts)
        {
            var relId = rel.RelationshipId;

            // Skip well-known structural relationships (slide layout, notes, theme, etc.)
            if (rel.OpenXmlPart is SlideLayoutPart or NotesSlidePart or ThemePart or ChartPart)
                continue;

            if (!referencedRelIds.Contains(relId))
            {
                var partType = rel.OpenXmlPart.GetType().Name;
                issues.Add(new ValidationIssue(
                    SlideNumber: slideNumber,
                    Severity: ValidationSeverity.Warning,
                    Category: "OrphanedRelationship",
                    Description: $"Slide {slideNumber}: Part '{relId}' ({partType}) is not referenced in the slide XML.",
                    Recommendation: "This part may be left over from a deleted shape. It wastes file space but is not harmful.",
                    XmlContext: relId));
            }
        }
    }

    private static void CheckHyperlinkTargets(SlidePart slidePart, int slideNumber, List<ValidationIssue> issues)
    {
        foreach (var hlink in slidePart.Slide.Descendants<A.HyperlinkOnClick>())
        {
            var relId = hlink.Id?.Value;
            if (string.IsNullOrEmpty(relId)) continue;

            // Internal slide links use action attribute
            var action = hlink.Action?.Value;
            if (action is not null && action.Contains("hlinksldjump", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var targetPart = slidePart.GetPartById(relId);
                    if (targetPart is not SlidePart)
                    {
                        issues.Add(new ValidationIssue(
                            SlideNumber: slideNumber,
                            Severity: ValidationSeverity.Error,
                            Category: "BrokenHyperlinkTarget",
                            Description: $"Slide {slideNumber}: Internal slide link '{relId}' does not point to a valid slide.",
                            Recommendation: "Update or remove the hyperlink. The link will fail when clicked in PowerPoint.",
                            XmlContext: $"a:hlinkClick[@r:id='{relId}']"));
                    }
                }
                catch
                {
                    issues.Add(new ValidationIssue(
                        SlideNumber: slideNumber,
                        Severity: ValidationSeverity.Error,
                        Category: "BrokenHyperlinkTarget",
                        Description: $"Slide {slideNumber}: Internal slide link '{relId}' references a missing part.",
                        Recommendation: "Remove the broken hyperlink. The target slide no longer exists.",
                        XmlContext: $"a:hlinkClick[@r:id='{relId}']"));
                }
                continue;
            }

            // External hyperlinks
            var hyperlinkRel = slidePart.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            if (hyperlinkRel is null)
            {
                issues.Add(new ValidationIssue(
                    SlideNumber: slideNumber,
                    Severity: ValidationSeverity.Warning,
                    Category: "BrokenHyperlinkTarget",
                    Description: $"Slide {slideNumber}: Hyperlink references relationship '{relId}' which does not exist.",
                    Recommendation: "Remove or re-create the hyperlink. It will not work when clicked.",
                    XmlContext: $"a:hlinkClick[@r:id='{relId}']"));
            }
        }
    }

    /// <summary>
    /// Tries to parse the XML of each related part to detect corruption.
    /// The OpenXML SDK loads part XML lazily, so this forces parsing without touching the slide XML.
    /// </summary>
    private static void CheckCorruptPartXml(SlidePart slidePart, int slideNumber, List<ValidationIssue> issues)
    {
        List<DocumentFormat.OpenXml.Packaging.IdPartPair> parts;
        try
        {
            parts = slidePart.Parts.ToList();
        }
        catch (XmlException ex)
        {
            issues.Add(new ValidationIssue(
                SlideNumber: slideNumber,
                Severity: ValidationSeverity.Error,
                Category: "CorruptRelationshipXml",
                Description: $"Slide {slideNumber}: The slide's relationship file (.rels) contains malformed XML.",
                Recommendation: "The relationship file is corrupt. This slide may not load or render correctly.",
                XmlContext: $"Line {ex.LineNumber}, Position {ex.LinePosition}"));
            return;
        }

        foreach (var partRef in parts)
        {
            // Skip structural parts accessed separately (slide XML handled by the outer try/catch)
            if (partRef.OpenXmlPart is SlideLayoutPart or SlideMasterPart or NotesSlidePart or ThemePart or SlidePart)
                continue;

            try
            {
                // Force lazy XML parsing — throws XmlException if the part XML is malformed
                var _ = partRef.OpenXmlPart.RootElement;
            }
            catch (XmlException ex)
            {
                var partType = partRef.OpenXmlPart.GetType().Name;
                issues.Add(new ValidationIssue(
                    SlideNumber: slideNumber,
                    Severity: ValidationSeverity.Error,
                    Category: "CorruptPartXml",
                    Description: $"Slide {slideNumber}: Related part '{partRef.RelationshipId}' ({partType}) contains malformed XML.",
                    Recommendation: "The part XML is corrupt. Consider recreating the affected content (e.g. re-insert the chart or image).",
                    XmlContext: $"{partRef.RelationshipId}, Line {ex.LineNumber}, Position {ex.LinePosition}"));
            }
        }
    }

    private static void CheckCrossSlideShapeIdDuplicates(
        PresentationDocument doc,
        IReadOnlyList<SlideId> slideIds,
        List<ValidationIssue> issues)
    {
        // Track (shapeId -> first slide number) across all slides
        var globalIds = new Dictionary<uint, int>();

        for (int i = 0; i < slideIds.Count; i++)
        {
            int currentSlideNumber = i + 1;
            var slidePart = GetSlidePart(doc, slideIds, i);

            ShapeTree? shapeTree;
            try
            {
                shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
            }
            catch (XmlException)
            {
                // Corrupt slide XML — already reported in the per-slide loop; skip here.
                continue;
            }

            if (shapeTree is null) continue;

            foreach (var child in shapeTree.ChildElements)
            {
                var (id, name) = GetShapeIdAndName(child);
                if (id is null) continue;

                if (globalIds.TryGetValue(id.Value, out var firstSlide))
                {
                    // Only report as info — cross-slide duplicates are common in normal presentations
                    issues.Add(new ValidationIssue(
                        SlideNumber: currentSlideNumber,
                        Severity: ValidationSeverity.Info,
                        Category: "CrossSlideDuplicateShapeId",
                        Description: $"Shape ID {id.Value} ('{name}') on slide {currentSlideNumber} also appears on slide {firstSlide}.",
                        Recommendation: "Cross-slide shape ID reuse is common and usually harmless, but may cause issues with some automation tools.",
                        XmlContext: $"p:sld/p:cSld/p:spTree/{child.LocalName}[cNvPr[@id='{id}']]"));
                }
                else
                {
                    globalIds[id.Value] = currentSlideNumber;
                }
            }
        }
    }

    private static (uint? Id, string Name) GetShapeIdAndName(DocumentFormat.OpenXml.OpenXmlElement element)
    {
        P.NonVisualDrawingProperties? nvProps = element switch
        {
            Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties,
            P.Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties,
            P.GraphicFrame gf => gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
            P.GroupShape gs => gs.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
            P.ConnectionShape cs => cs.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties,
            _ => null
        };

        if (nvProps is null) return (null, "(unknown)");

        return (nvProps.Id?.Value, nvProps.Name?.Value ?? "(unnamed)");
    }
}
