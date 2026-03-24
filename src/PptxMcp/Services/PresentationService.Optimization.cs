using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using PptxMcp.Models;

namespace PptxMcp.Services;

public partial class PresentationService
{
    /// <summary>
    /// Analyze the file size breakdown of a PPTX file by category.
    /// Opens the package read-only via System.IO.Packaging and enumerates all parts.
    /// </summary>
    public FileSizeAnalysisResult AnalyzeFileSize(string filePath)
    {
        var fileInfo = new FileInfo(filePath);
        var totalFileSize = fileInfo.Length;

        using var package = Package.Open(filePath, FileMode.Open, FileAccess.Read);

        // Uses OPC Package.GetParts() which enumerates logical parts (not raw ZIP entries).
        // This is correct for categorization; raw ZIP sizes would require ZipArchive separately.
        var categoryParts = new Dictionary<string, List<FileSizePart>>
        {
            ["slides"] = [],
            ["images"] = [],
            ["video_audio"] = [],
            ["masters"] = [],
            ["layouts"] = [],
            ["other"] = [],
        };

        foreach (var part in package.GetParts())
        {
            var uri = part.Uri.ToString();
            var contentType = part.ContentType;
            long size;
            using (var stream = part.GetStream(FileMode.Open, FileAccess.Read))
            {
                size = stream.Length;
            }

            var category = CategorizePart(uri, contentType);
            categoryParts[category].Add(new FileSizePart(uri, contentType, size));
        }

        var categories = categoryParts
            .Select(kvp => new FileSizeCategory(
                Name: kvp.Key,
                TotalSize: kvp.Value.Sum(p => p.Size),
                PartCount: kvp.Value.Count,
                Parts: kvp.Value))
            .ToList();

        var totalPartSize = categories.Sum(c => c.TotalSize);

        return new FileSizeAnalysisResult(
            Success: true,
            FilePath: filePath,
            TotalFileSize: totalFileSize,
            TotalPartSize: totalPartSize,
            Categories: categories,
            Message: $"Analyzed {categories.Sum(c => c.PartCount)} parts across {categories.Count} categories.");
    }

    private static string CategorizePart(string uri, string contentType)
    {
        var lowerUri = uri.ToLowerInvariant();

        // Slide XML
        if (lowerUri.StartsWith("/ppt/slides/") && !lowerUri.EndsWith(".rels"))
            return "slides";

        // Slide masters
        if (lowerUri.StartsWith("/ppt/slidemasters/") && !lowerUri.EndsWith(".rels"))
            return "masters";

        // Slide layouts
        if (lowerUri.StartsWith("/ppt/slidelayouts/") && !lowerUri.EndsWith(".rels"))
            return "layouts";

        // Media — images
        if (contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
            return "images";

        // Media — video/audio
        if (contentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase) ||
            contentType.StartsWith("audio/", StringComparison.OrdinalIgnoreCase))
            return "video_audio";

        // Media folder fallback (catches media with unusual content types)
        if (lowerUri.StartsWith("/ppt/media/"))
        {
            if (contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                return "images";
            return "video_audio";
        }

        return "other";
    }

    public UnusedLayoutsResult FindUnusedLayouts(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation part is missing.");

        // Build a map from each SlideLayoutPart URI to the 1-based slide numbers that reference it.
        var layoutUsage = new Dictionary<string, List<int>>();
        int slideNumber = 0;
        foreach (var slidePart in GetOrderedSlideParts(presentationPart))
        {
            slideNumber++;
            if (slidePart.SlideLayoutPart is { } layoutPart)
            {
                var uri = layoutPart.Uri.ToString();
                if (!layoutUsage.TryGetValue(uri, out var list))
                {
                    list = [];
                    layoutUsage[uri] = list;
                }
                list.Add(slideNumber);
            }
        }

        var masters = new List<MasterInfo>();
        var layouts = new List<LayoutInfo>();
        var warnings = new List<string>();
        long estimatedSavings = 0;

        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            var masterName = masterPart.SlideMaster?.CommonSlideData?.Name?.Value ?? "Unnamed Master";
            var masterUri = masterPart.Uri.ToString();
            long masterSize = EstimatePartSize(masterPart);

            int layoutCount = 0;
            int usedLayoutCount = 0;

            foreach (var layoutPart in masterPart.SlideLayoutParts)
            {
                layoutCount++;
                var layoutName = layoutPart.SlideLayout?.CommonSlideData?.Name?.Value ?? "Unnamed Layout";
                var layoutUri = layoutPart.Uri.ToString();
                long layoutSize = EstimatePartSize(layoutPart);
                bool layoutUsed = layoutUsage.ContainsKey(layoutUri);

                if (layoutUsed)
                    usedLayoutCount++;
                else
                    estimatedSavings += layoutSize;

                layouts.Add(new LayoutInfo(
                    Name: layoutName,
                    Uri: layoutUri,
                    SizeBytes: layoutSize,
                    IsUsed: layoutUsed,
                    MasterName: masterName,
                    ReferencedBySlides: layoutUsed ? layoutUsage[layoutUri].AsReadOnly() : []));
            }

            bool masterUsed = usedLayoutCount > 0;
            if (!masterUsed)
                estimatedSavings += masterSize;

            masters.Add(new MasterInfo(
                Name: masterName,
                Uri: masterUri,
                SizeBytes: masterSize,
                IsUsed: masterUsed,
                LayoutCount: layoutCount,
                UsedLayoutCount: usedLayoutCount));

            // Warn if removing master would orphan layouts that are still in use.
            if (!masterUsed && usedLayoutCount > 0)
            {
                warnings.Add(
                    $"Removing master '{masterName}' would orphan {usedLayoutCount} layout(s) still in use.");
            }
        }

        int unusedMasterCount = masters.Count(m => !m.IsUsed);
        int unusedLayoutCount = layouts.Count(l => !l.IsUsed);

        string message = unusedLayoutCount == 0 && unusedMasterCount == 0
            ? "All masters and layouts are in use."
            : $"Found {unusedLayoutCount} unused layout(s) and {unusedMasterCount} unused master(s). " +
              $"Estimated savings: {estimatedSavings:N0} bytes.";

        return new UnusedLayoutsResult(
            Success: true,
            FilePath: filePath,
            TotalMasters: masters.Count,
            TotalLayouts: layouts.Count,
            UnusedMasterCount: unusedMasterCount,
            UnusedLayoutCount: unusedLayoutCount,
            EstimatedSavingsBytes: estimatedSavings,
            Masters: masters,
            Layouts: layouts,
            Warnings: warnings,
            Message: message);
    }

    /// <summary>
    /// Returns slide parts in presentation order (matching SlideIdList).
    /// </summary>
    private static IEnumerable<SlidePart> GetOrderedSlideParts(PresentationPart presentationPart)
    {
        var slideIdList = presentationPart.Presentation.SlideIdList;
        if (slideIdList is null)
            yield break;

        foreach (var slideId in slideIdList.Elements<DocumentFormat.OpenXml.Presentation.SlideId>())
        {
            if (slideId.RelationshipId?.Value is { } relId &&
                presentationPart.GetPartById(relId) is SlidePart slidePart)
            {
                yield return slidePart;
            }
        }
    }

    /// <summary>
    /// Estimates the size of an OpenXML part by reading its stream length,
    /// plus the streams of any direct relationship parts (rels, theme, images, etc.).
    /// </summary>
    private static long EstimatePartSize(OpenXmlPart part)
    {
        long size = 0;
        try
        {
            using var stream = part.GetStream();
            size += stream.Length;
        }
        catch
        {
            // Part stream may be inaccessible; skip.
        }

        foreach (var rel in part.Parts)
        {
            try
            {
                using var relStream = rel.OpenXmlPart.GetStream();
                size += relStream.Length;
            }
            catch
            {
                // Skip inaccessible relationship parts.
            }
        }

        return size;
    }
}
