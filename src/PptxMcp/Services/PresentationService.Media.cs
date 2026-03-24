using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxMcp.Models;

namespace PptxMcp.Services;

public partial class PresentationService
{
    public MediaAnalysisResult AnalyzeMedia(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var presentationPart = doc.PresentationPart;
        if (presentationPart is null)
            return new MediaAnalysisResult(
                Success: false, FilePath: filePath,
                TotalMediaCount: 0, TotalMediaSize: 0,
                DuplicateGroupCount: 0, DuplicateSavingsBytes: 0,
                MediaParts: [], DuplicateGroups: [],
                Message: "Presentation part not found.");

        var slideNumberByPartUri = BuildMediaSlideNumberLookup(presentationPart);
        var mediaInfoMap = new Dictionary<string, MediaPartBuilder>();

        // Collect media from each slide
        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<SlideId>() ?? [];
        foreach (var slideId in slideIds)
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!);
            var slideNumber = slideNumberByPartUri[slidePart.Uri.ToString()];
            CollectMediaFromOpenXmlPart(slidePart, slideNumber, mediaInfoMap);
        }

        // Scan slide layouts and masters for inherited media
        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            CollectMediaFromOpenXmlPart(masterPart, slideNumber: 0, mediaInfoMap);
            foreach (var layoutPart in masterPart.SlideLayoutParts)
                CollectMediaFromOpenXmlPart(layoutPart, slideNumber: 0, mediaInfoMap);
        }

        if (mediaInfoMap.Count == 0)
            return new MediaAnalysisResult(
                Success: true, FilePath: filePath,
                TotalMediaCount: 0, TotalMediaSize: 0,
                DuplicateGroupCount: 0, DuplicateSavingsBytes: 0,
                MediaParts: [], DuplicateGroups: [],
                Message: "No media assets found in the presentation.");

        var mediaParts = mediaInfoMap.Values
            .OrderBy(m => m.Path)
            .Select(m => new MediaPartInfo(
                Path: m.Path,
                ContentType: m.ContentType,
                SizeBytes: m.SizeBytes,
                Hash: m.Hash,
                ReferencedBySlides: m.SlideNumbers.OrderBy(n => n).Distinct().ToArray()))
            .ToArray();

        var totalSize = mediaParts.Sum(p => p.SizeBytes);

        var duplicateGroups = mediaParts
            .GroupBy(p => p.Hash)
            .Where(g => g.Count() > 1)
            .Select(g => new DuplicateGroup(
                Hash: g.Key,
                ContentType: g.First().ContentType,
                SizeBytes: g.First().SizeBytes,
                Parts: g.Select(p => p.Path).ToArray(),
                ReferencedBySlides: g.SelectMany(p => p.ReferencedBySlides).Distinct().OrderBy(n => n).ToArray()))
            .ToArray();

        var duplicateSavings = duplicateGroups.Sum(g => g.SizeBytes * (g.Parts.Length - 1));

        var message = duplicateGroups.Length > 0
            ? $"Found {mediaParts.Length} media asset(s) with {duplicateGroups.Length} duplicate group(s). Potential savings: {duplicateSavings:N0} bytes."
            : $"Found {mediaParts.Length} media asset(s). No duplicates detected.";

        return new MediaAnalysisResult(
            Success: true,
            FilePath: filePath,
            TotalMediaCount: mediaParts.Length,
            TotalMediaSize: totalSize,
            DuplicateGroupCount: duplicateGroups.Length,
            DuplicateSavingsBytes: duplicateSavings,
            MediaParts: mediaParts,
            DuplicateGroups: duplicateGroups,
            Message: message);
    }

    private static Dictionary<string, int> BuildMediaSlideNumberLookup(PresentationPart presentationPart)
    {
        var lookup = new Dictionary<string, int>();
        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<SlideId>() ?? [];
        int slideNumber = 1;
        foreach (var slideId in slideIds)
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!);
            lookup[slidePart.Uri.ToString()] = slideNumber++;
        }
        return lookup;
    }

    /// <summary>Collect image parts (OpenXmlPart) and video/audio parts (DataPart) from a part.</summary>
    private static void CollectMediaFromOpenXmlPart(OpenXmlPart ownerPart, int slideNumber, Dictionary<string, MediaPartBuilder> mediaMap)
    {
        // ImagePart inherits from OpenXmlPart — accessible via Parts collection
        foreach (var idPartPair in ownerPart.Parts)
        {
            if (idPartPair.OpenXmlPart is ImagePart imagePart)
                RegisterMediaPart(imagePart.Uri.ToString(), imagePart.ContentType, () => imagePart.GetStream(), slideNumber, mediaMap);
        }

        // MediaDataPart inherits from DataPart — accessible via DataPartReferenceRelationships
        foreach (var dataPartRef in ownerPart.DataPartReferenceRelationships)
        {
            var dataPart = dataPartRef.DataPart;
            RegisterMediaPart(dataPart.Uri.ToString(), dataPart.ContentType, () => dataPart.GetStream(), slideNumber, mediaMap);
        }
    }

    private static void RegisterMediaPart(string uri, string contentType, Func<Stream> getStream, int slideNumber, Dictionary<string, MediaPartBuilder> mediaMap)
    {
        if (!mediaMap.TryGetValue(uri, out var builder))
        {
            using var stream = getStream();
            var size = stream.Length;
            stream.Position = 0;
            var hash = Convert.ToHexString(SHA256.HashData(stream));

            builder = new MediaPartBuilder
            {
                Path = uri,
                ContentType = contentType,
                SizeBytes = size,
                Hash = hash
            };
            mediaMap[uri] = builder;
        }

        if (slideNumber > 0)
            builder.SlideNumbers.Add(slideNumber);
    }

    private sealed class MediaPartBuilder
    {
        public required string Path { get; init; }
        public required string ContentType { get; init; }
        public required long SizeBytes { get; init; }
        public required string Hash { get; init; }
        public HashSet<int> SlideNumbers { get; } = [];
    }
}
