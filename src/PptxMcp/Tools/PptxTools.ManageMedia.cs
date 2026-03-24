using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxMcp.Models;

namespace PptxMcp.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Manage media assets in a PowerPoint presentation.
    /// Available actions:
    /// - Analyze: List all media assets (images, video, audio), detect duplicates by SHA256 hash (read-only).
    /// - Deduplicate: Consolidate identical media, redirect references, and remove orphaned copies.
    /// - AnalyzeVideo: Extract video/audio metadata (codec, resolution, duration, bitrate) from embedded media (read-only).
    /// Natural workflow: Analyze (identify duplicates) → Deduplicate (clean them up). Use AnalyzeVideo for media inspection.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="action">The media management operation to perform: Analyze, Deduplicate, or AnalyzeVideo.</param>
    [McpServerTool(Title = "Manage Media")]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["Analyze","Deduplicate","AnalyzeVideo"]""")]
    public partial Task<string> pptx_manage_media(
        string filePath,
        ManageMediaAction action)
    {
        return action switch
        {
            ManageMediaAction.Analyze => ExecuteToolStructured(filePath,
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
                    Message: error)),

            ManageMediaAction.Deduplicate => ExecuteToolStructured(filePath,
                () => _service.DeduplicateMedia(filePath),
                error => new DeduplicateMediaResult(
                    Success: false,
                    FilePath: filePath,
                    DuplicateGroupsFound: 0,
                    PartsRemoved: 0,
                    BytesSaved: 0,
                    Groups: [],
                    Validation: new ValidationStatus(0, 0, false),
                    Message: error)),

            ManageMediaAction.AnalyzeVideo => ExecuteToolStructured(filePath,
                () => _service.AnalyzeVideoMetadata(filePath),
                error => new VideoMetadataResult(
                    Success: false,
                    FilePath: filePath,
                    VideoPartsFound: 0,
                    TotalTracks: 0,
                    Parts: [],
                    Message: error)),

            _ => Task.FromResult(JsonSerializer.Serialize(
                new { Success = false, Message = $"Unknown action: {action}. Valid actions: Analyze, Deduplicate, AnalyzeVideo." },
                IndentedJson))
        };
    }
}
