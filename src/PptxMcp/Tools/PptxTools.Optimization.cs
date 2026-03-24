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
                Categories: [],
                Message: error));
}
