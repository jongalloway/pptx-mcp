using ModelContextProtocol.Server;
using PptxMcp.Models;

namespace PptxMcp.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Deduplicate identical media in a PowerPoint presentation.
    /// Finds media parts with the same content (SHA256 hash match), redirects all references
    /// to a single canonical copy, and removes orphaned duplicates.
    /// Validates the package with OpenXmlValidator before and after modification.
    /// Returns structured JSON with deduplication statistics and space saved.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file to modify.</param>
    [McpServerTool(Title = "Deduplicate Media")]
    public partial Task<string> pptx_deduplicate_media(string filePath) =>
        ExecuteToolStructured(filePath,
            () => _service.DeduplicateMedia(filePath),
            error => new DeduplicateMediaResult(
                Success: false,
                FilePath: filePath,
                DuplicateGroupsFound: 0,
                PartsRemoved: 0,
                BytesSaved: 0,
                Groups: [],
                Validation: new ValidationStatus(0, 0, false),
                Message: error));
}
