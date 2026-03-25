namespace PptxTools.Models;

/// <summary>Actions for the consolidated pptx_manage_media tool.</summary>
public enum ManageMediaAction
{
    /// <summary>Analyze all media assets, detecting duplicates by SHA256 hash.</summary>
    Analyze,

    /// <summary>Deduplicate identical media, redirecting references and removing orphans.</summary>
    Deduplicate,

    /// <summary>Extract video/audio metadata (codec, resolution, duration, bitrate) from embedded media.</summary>
    AnalyzeVideo
}
