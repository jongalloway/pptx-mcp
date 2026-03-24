using ModelContextProtocol.Server;
using PptxMcp.Models;

namespace PptxMcp.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Analyze the file size breakdown of a PowerPoint presentation by category.
    /// Scans all parts in the PPTX package (ZIP structure) and reports sizes broken down into:
    /// slides, images, video/audio, slide masters, slide layouts, and other parts.
    /// Each category includes a subtotal and per-part detail (relative path, content type, size in bytes).
    /// The root level includes actual file size on disk and total uncompressed part size.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    [McpServerTool(Title = "Analyze File Size", ReadOnly = true, Idempotent = true)]
    public partial Task<string> pptx_analyze_file_size(string filePath) =>
        ExecuteToolStructured(filePath,
            () => _service.AnalyzeFileSize(filePath),
            error => new FileSizeAnalysisResult(
                Success: false,
                FilePath: filePath,
                TotalFileSize: 0,
                TotalPartSize: 0,
                Categories: EmptyFileSizeCategories,
                Message: error));

    private static readonly IReadOnlyList<FileSizeCategory> EmptyFileSizeCategories =
    [
        new("slides", 0, 0, []),
        new("images", 0, 0, []),
        new("video_audio", 0, 0, []),
        new("masters", 0, 0, []),
        new("layouts", 0, 0, []),
        new("other", 0, 0, []),
    ];

    /// <summary>
    /// Find unused slide masters and layouts in a PowerPoint presentation.
    /// Enumerates all masters and layouts, cross-references against actual slide usage,
    /// and identifies which could be safely removed with estimated space savings.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    [McpServerTool(Title = "Find Unused Layouts", ReadOnly = true, Idempotent = true)]
    public partial Task<string> pptx_find_unused_layouts(string filePath) =>
        ExecuteToolStructured(filePath,
            () => _service.FindUnusedLayouts(filePath),
            error => new UnusedLayoutsResult(
                Success: false,
                FilePath: filePath,
                TotalMasters: 0,
                TotalLayouts: 0,
                UnusedMasterCount: 0,
                UnusedLayoutCount: 0,
                EstimatedSavingsBytes: 0,
                Masters: [],
                Layouts: [],
                Warnings: [],
                Message: error));

    /// <summary>
    /// Remove unused slide layouts and orphaned slide masters from a PowerPoint presentation.
    /// When layoutUris is omitted, auto-detects and removes all unused layouts.
    /// When specific URIs are provided, removes only those (if they are unused).
    /// Validates the package with OpenXmlValidator before and after removal.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file to modify.</param>
    /// <param name="layoutUris">Optional array of layout URIs to remove. Omit to auto-detect all unused layouts.</param>
    [McpServerTool(Title = "Remove Unused Layouts")]
    public partial Task<string> pptx_remove_unused_layouts(string filePath, string[]? layoutUris = null) =>
        ExecuteToolStructured(filePath,
            () => _service.RemoveUnusedLayouts(filePath, layoutUris),
            error => new RemoveLayoutsResult(
                Success: false,
                FilePath: filePath,
                RemovedItems: [],
                LayoutsRemoved: 0,
                MastersRemoved: 0,
                BytesSaved: 0,
                Validation: new ValidationStatus(0, 0, false),
                Message: error));

    /// <summary>
    /// Optimize images in a PowerPoint presentation by downscaling, converting formats, and recompressing.
    /// Scans all images across slides, layouts, and masters. Downscales images that are larger than their
    /// display dimensions warrant based on target DPI. Converts BMP/TIFF to PNG/JPEG. Recompresses JPEG images
    /// at the specified quality level. Only replaces images when optimization results in smaller file size.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file to modify.</param>
    /// <param name="targetDpi">Target DPI for screen display (default 150; use 300 for print).</param>
    /// <param name="jpegQuality">JPEG compression quality 1-100 (default 85; higher = larger file).</param>
    /// <param name="convertFormats">Convert BMP/TIFF to PNG/JPEG (default true).</param>
    [McpServerTool(Title = "Optimize Images")]
    public partial Task<string> pptx_optimize_images(
        string filePath,
        int targetDpi = 150,
        int jpegQuality = 85,
        bool convertFormats = true) =>
        ExecuteToolStructured(filePath,
            () => _service.OptimizeImages(filePath, targetDpi, jpegQuality, convertFormats),
            error => new ImageOptimizationResult(
                Success: false,
                FilePath: filePath,
                ImagesProcessed: 0,
                ImagesSkipped: 0,
                TotalBytesBefore: 0,
                TotalBytesAfter: 0,
                TotalBytesSaved: 0,
                OptimizedImages: [],
                Validation: new ValidationStatus(0, 0, false),
                Message: error));
}
