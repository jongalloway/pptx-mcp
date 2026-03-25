namespace PptxTools.Models;

/// <summary>Details about a single optimized image.</summary>
/// <param name="ImagePath">Package URI of the image part.</param>
/// <param name="OriginalFormat">Original image format (PNG, JPEG, BMP, TIFF, etc.).</param>
/// <param name="OptimizedFormat">Format after optimization.</param>
/// <param name="OriginalWidth">Original image width in pixels.</param>
/// <param name="OriginalHeight">Original image height in pixels.</param>
/// <param name="OptimizedWidth">Width after downscaling (same as original if not downscaled).</param>
/// <param name="OptimizedHeight">Height after downscaling (same as original if not downscaled).</param>
/// <param name="OriginalSizeBytes">Original image size in bytes.</param>
/// <param name="OptimizedSizeBytes">Size after optimization in bytes.</param>
/// <param name="BytesSaved">Bytes saved by optimization.</param>
/// <param name="Action">Description of action taken (downscaled, converted, recompressed, skipped).</param>
public record OptimizedImageInfo(
    string ImagePath,
    string OriginalFormat,
    string OptimizedFormat,
    int OriginalWidth,
    int OriginalHeight,
    int OptimizedWidth,
    int OptimizedHeight,
    long OriginalSizeBytes,
    long OptimizedSizeBytes,
    long BytesSaved,
    string Action);

/// <summary>Structured result for pptx_optimize_images.</summary>
/// <param name="Success">True when the operation completed without errors.</param>
/// <param name="FilePath">Path to the modified presentation file.</param>
/// <param name="ImagesProcessed">Number of images that were optimized.</param>
/// <param name="ImagesSkipped">Number of images that were skipped (no optimization possible).</param>
/// <param name="TotalBytesBefore">Total size of all images before optimization.</param>
/// <param name="TotalBytesAfter">Total size of all images after optimization.</param>
/// <param name="TotalBytesSaved">Total bytes saved by optimization.</param>
/// <param name="OptimizedImages">Details for each optimized or skipped image.</param>
/// <param name="Validation">OpenXML validation status before and after.</param>
/// <param name="Message">Human-readable status or error message.</param>
public record ImageOptimizationResult(
    bool Success,
    string FilePath,
    int ImagesProcessed,
    int ImagesSkipped,
    long TotalBytesBefore,
    long TotalBytesAfter,
    long TotalBytesSaved,
    IReadOnlyList<OptimizedImageInfo> OptimizedImages,
    ValidationStatus Validation,
    string Message);
