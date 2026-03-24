using ModelContextProtocol.Server;
using PptxMcp.Models;

namespace PptxMcp.Tools;

public partial class PptxTools
{
    /// <summary>
    /// List and analyze all media assets (images, video, audio) in a PowerPoint presentation.
    /// For each media part: reports name, content type, size, SHA256 hash, and which slides reference it.
    /// Detects duplicate media (same content hash) and groups them for deduplication analysis.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    [McpServerTool(Title = "Analyze Media Assets", ReadOnly = true, Idempotent = true)]
    public partial Task<string> pptx_analyze_media(string filePath) =>
        ExecuteToolStructured(filePath,
            () => _service.AnalyzeMedia(filePath),
            error => new MediaAnalysisResult(
                Success: false,
                FilePath: filePath,
                TotalMediaCount: 0,
                TotalMediaSize: 0,
                DuplicateGroupCount: 0,
                DuplicateSavingsBytes: 0,
                MediaParts: [],
                DuplicateGroups: [],
                Message: error));
}
