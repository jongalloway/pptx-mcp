namespace PptxTools.Models;

/// <summary>Structured result for pptx_analyze_file_size.</summary>
/// <param name="Success">True when analysis completed successfully.</param>
/// <param name="FilePath">Path to the analyzed PPTX file.</param>
/// <param name="TotalFileSize">Actual file size on disk in bytes.</param>
/// <param name="TotalPartSize">Sum of all uncompressed part sizes in bytes.</param>
/// <param name="Categories">Breakdown by category with subtotals and per-part detail.</param>
/// <param name="Message">Human-readable status or error message.</param>
public record FileSizeAnalysisResult(
    bool Success,
    string FilePath,
    long TotalFileSize,
    long TotalPartSize,
    IReadOnlyList<FileSizeCategory> Categories,
    string Message);

/// <summary>A single category in the file size breakdown.</summary>
/// <param name="Name">Category name (slides, images, video_audio, masters, layouts, other).</param>
/// <param name="TotalSize">Sum of uncompressed sizes for all parts in this category.</param>
/// <param name="PartCount">Number of parts in this category.</param>
/// <param name="Parts">Individual parts with path, content type, and size.</param>
public record FileSizeCategory(
    string Name,
    long TotalSize,
    int PartCount,
    IReadOnlyList<FileSizePart> Parts);

/// <summary>Size information for a single part in the PPTX package.</summary>
/// <param name="Path">Relative path within the package (e.g. /ppt/slides/slide1.xml).</param>
/// <param name="ContentType">MIME content type of the part.</param>
/// <param name="Size">Uncompressed size in bytes.</param>
public record FileSizePart(
    string Path,
    string ContentType,
    long Size);
