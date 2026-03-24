namespace PptxMcp.Models;

/// <summary>Structured result for pptx_analyze_media.</summary>
/// <param name="Success">True when media analysis completed successfully.</param>
/// <param name="FilePath">Path to the analyzed presentation file.</param>
/// <param name="TotalMediaCount">Total number of distinct media parts in the package.</param>
/// <param name="TotalMediaSize">Total size of all media parts in bytes.</param>
/// <param name="DuplicateGroupCount">Number of groups containing identical media (same SHA256 hash).</param>
/// <param name="DuplicateSavingsBytes">Bytes that could be saved by deduplicating identical media.</param>
/// <param name="MediaParts">List of all media parts with metadata.</param>
/// <param name="DuplicateGroups">Groups of media parts that share the same content hash.</param>
/// <param name="Message">Human-readable status message describing the outcome.</param>
public record MediaAnalysisResult(
    bool Success,
    string? FilePath,
    int TotalMediaCount,
    long TotalMediaSize,
    int DuplicateGroupCount,
    long DuplicateSavingsBytes,
    MediaPartInfo[] MediaParts,
    DuplicateGroup[] DuplicateGroups,
    string Message);

/// <summary>Metadata for a single media part in the presentation package.</summary>
/// <param name="Path">URI path of the media part within the package (e.g. /ppt/media/image1.png).</param>
/// <param name="ContentType">MIME content type (e.g. image/png, video/mp4).</param>
/// <param name="SizeBytes">Size of the media data in bytes.</param>
/// <param name="Hash">SHA256 hex hash of the media content.</param>
/// <param name="ReferencedBySlides">1-based slide numbers that reference this media part.</param>
public record MediaPartInfo(
    string Path,
    string ContentType,
    long SizeBytes,
    string Hash,
    int[] ReferencedBySlides);

/// <summary>A group of media parts that share identical content (same SHA256 hash).</summary>
/// <param name="Hash">SHA256 hex hash shared by all parts in this group.</param>
/// <param name="ContentType">MIME content type of the duplicated media.</param>
/// <param name="SizeBytes">Size of each individual copy in bytes.</param>
/// <param name="Parts">URI paths of all parts sharing this hash.</param>
/// <param name="ReferencedBySlides">Combined 1-based slide numbers that reference any part in this group.</param>
public record DuplicateGroup(
    string Hash,
    string ContentType,
    long SizeBytes,
    string[] Parts,
    int[] ReferencedBySlides);
