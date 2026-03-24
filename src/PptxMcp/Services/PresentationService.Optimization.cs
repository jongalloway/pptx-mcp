using System.IO.Packaging;
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

        var categoryParts = new Dictionary<string, List<FileSizePart>>
        {
            ["slides"] = [],
            ["images"] = [],
            ["videoAudio"] = [],
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
            return "videoAudio";

        // Media folder fallback (catches media with unusual content types)
        if (lowerUri.StartsWith("/ppt/media/"))
        {
            if (contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                return "images";
            return "videoAudio";
        }

        return "other";
    }
}
