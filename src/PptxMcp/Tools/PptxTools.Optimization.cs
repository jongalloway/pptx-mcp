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
}
