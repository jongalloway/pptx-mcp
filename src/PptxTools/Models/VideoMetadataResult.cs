namespace PptxTools.Models;

/// <summary>Metadata extracted from a single video or audio track within an embedded media part.</summary>
/// <param name="TrackType">Track type: "video" or "audio".</param>
/// <param name="Codec">Codec four-character code (e.g. "avc1", "hev1", "mp4a").</param>
/// <param name="Width">Video width in pixels (null for audio tracks).</param>
/// <param name="Height">Video height in pixels (null for audio tracks).</param>
/// <param name="DurationSeconds">Track duration in seconds (null if unavailable).</param>
/// <param name="Bitrate">Average bitrate in bits per second (null if unavailable).</param>
/// <param name="ChannelCount">Audio channel count (null for video tracks).</param>
/// <param name="SampleRate">Audio sample rate in Hz (null for video tracks).</param>
public record VideoTrackInfo(
    string TrackType,
    string Codec,
    int? Width,
    int? Height,
    double? DurationSeconds,
    long? Bitrate,
    int? ChannelCount,
    int? SampleRate);

/// <summary>Metadata for a single embedded media part in the presentation.</summary>
/// <param name="PartUri">Package URI of the media part (e.g. /ppt/media/video1.mp4).</param>
/// <param name="ContentType">MIME content type of the media part.</param>
/// <param name="FileSizeBytes">Size of the media data in bytes.</param>
/// <param name="Tracks">Parsed track metadata (video and/or audio tracks).</param>
/// <param name="Error">Error message if parsing failed for this part, otherwise null.</param>
public record VideoPartInfo(
    string PartUri,
    string ContentType,
    long FileSizeBytes,
    IReadOnlyList<VideoTrackInfo> Tracks,
    string? Error);

/// <summary>Structured result for the AnalyzeVideo action of pptx_manage_media.</summary>
/// <param name="Success">True when analysis completed without fatal errors.</param>
/// <param name="FilePath">Path to the analyzed presentation file.</param>
/// <param name="VideoPartsFound">Number of video/audio media parts found.</param>
/// <param name="TotalTracks">Total number of tracks extracted across all parts.</param>
/// <param name="Parts">Per-part metadata details.</param>
/// <param name="Message">Human-readable status or error message.</param>
public record VideoMetadataResult(
    bool Success,
    string FilePath,
    int VideoPartsFound,
    int TotalTracks,
    IReadOnlyList<VideoPartInfo> Parts,
    string Message);
