using DocumentFormat.OpenXml.Packaging;
using PptxTools.Models;
using SharpISOBMFF;
using SharpMP4;

namespace PptxTools.Services;

public partial class PresentationService
{
    /// <summary>
    /// Analyze embedded video and audio media in a PPTX file, extracting codec,
    /// resolution, duration, and bitrate metadata from MP4/M4V/MOV containers.
    /// </summary>
    /// <param name="filePath">Path to the PPTX file.</param>
    public VideoMetadataResult AnalyzeVideoMetadata(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var presentationPart = doc.PresentationPart;
        if (presentationPart is null)
            return new VideoMetadataResult(
                Success: false, FilePath: filePath,
                VideoPartsFound: 0, TotalTracks: 0, Parts: [],
                Message: "Presentation part not found.");

        var parts = new List<VideoPartInfo>();

        // Collect video/audio DataParts from all slides, layouts, and masters.
        var processedUris = new HashSet<string>();
        var allOwnerParts = CollectAllOwnerParts(presentationPart);

        foreach (var ownerPart in allOwnerParts)
        {
            foreach (var dataPartRef in ownerPart.DataPartReferenceRelationships)
            {
                var dataPart = dataPartRef.DataPart;
                var uri = dataPart.Uri.ToString();
                if (!processedUris.Add(uri))
                    continue;

                if (!IsVideoOrAudioContentType(dataPart.ContentType))
                    continue;

                var partInfo = ExtractMediaPartMetadata(dataPart);
                parts.Add(partInfo);
            }
        }

        if (parts.Count == 0)
            return new VideoMetadataResult(
                Success: true, FilePath: filePath,
                VideoPartsFound: 0, TotalTracks: 0, Parts: [],
                Message: "No video or audio media found in the presentation.");

        int totalTracks = parts.Sum(p => p.Tracks.Count);
        int partsWithErrors = parts.Count(p => p.Error is not null);

        string message = partsWithErrors > 0
            ? $"Found {parts.Count} media part(s) with {totalTracks} track(s). {partsWithErrors} part(s) had parse errors."
            : $"Found {parts.Count} media part(s) with {totalTracks} track(s).";

        return new VideoMetadataResult(
            Success: true,
            FilePath: filePath,
            VideoPartsFound: parts.Count,
            TotalTracks: totalTracks,
            Parts: parts,
            Message: message);
    }

    private static bool IsVideoOrAudioContentType(string contentType)
    {
        return contentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase)
            || contentType.StartsWith("audio/", StringComparison.OrdinalIgnoreCase);
    }

    private static VideoPartInfo ExtractMediaPartMetadata(DataPart dataPart)
    {
        var uri = dataPart.Uri.ToString();
        var contentType = dataPart.ContentType;

        // Copy stream into memory so we can measure size and parse independently.
        var memoryStream = new MemoryStream();
        using (var raw = dataPart.GetStream())
        {
            raw.CopyTo(memoryStream);
        }
        long fileSize = memoryStream.Length;
        memoryStream.Position = 0;

        try
        {
            var tracks = ParseMp4Tracks(memoryStream, fileSize);
            return new VideoPartInfo(uri, contentType, fileSize, tracks, Error: null);
        }
        catch (Exception ex)
        {
            return new VideoPartInfo(uri, contentType, fileSize, [],
                Error: $"Failed to parse MP4 container: {ex.Message}");
        }
    }

    private static List<VideoTrackInfo> ParseMp4Tracks(MemoryStream stream, long fileSize)
    {
        var container = new Container();
        using var isoStream = new IsoStream(stream);
        container.Read(isoStream);

        var moov = Mp4Extensions.GetMovieBox(container);
        if (moov is null)
            return [];

        // Get movie-level timescale and duration from mvhd.
        var mvhd = moov.Children.OfType<MovieHeaderBox>().FirstOrDefault();
        uint movieTimescale = mvhd?.Timescale ?? 0;
        ulong movieDuration = mvhd?.Duration ?? 0;

        var trackBoxes = Mp4Extensions.GetTracks(moov);
        var results = new List<VideoTrackInfo>();

        foreach (var trackBox in trackBoxes)
        {
            var tkhd = trackBox.Children.OfType<TrackHeaderBox>().FirstOrDefault();
            var mdia = trackBox.Children.OfType<MediaBox>().FirstOrDefault();
            if (mdia is null) continue;

            var mdhd = mdia.Children.OfType<MediaHeaderBox>().FirstOrDefault();
            var hdlr = mdia.Children.OfType<HandlerBox>().FirstOrDefault();
            var minf = mdia.Children.OfType<MediaInformationBox>().FirstOrDefault();
            var stbl = minf?.Children.OfType<SampleTableBox>().FirstOrDefault();
            var stsd = stbl?.Children.OfType<SampleDescriptionBox>().FirstOrDefault();

            // Determine handler type.
            uint handlerType = hdlr?.HandlerType ?? 0;
            bool isVideo = IsHandlerType(handlerType, "vide");
            bool isAudio = IsHandlerType(handlerType, "soun");
            if (!isVideo && !isAudio) continue;

            // Duration from media header (mdhd), falling back to track header (tkhd) + movie timescale.
            double? durationSeconds = ComputeDuration(mdhd, tkhd, movieTimescale, movieDuration);

            // Bitrate: compute from file size and duration if only one track,
            // or from BitRateBox if available.
            long? bitrate = null;
            var bitrateBox = FindDescendant<BitRateBox>(stsd);
            if (bitrateBox is not null && bitrateBox.AvgBitrate > 0)
                bitrate = bitrateBox.AvgBitrate;
            else if (durationSeconds is > 0)
                bitrate = (long)(fileSize * 8.0 / durationSeconds.Value);

            if (isVideo)
            {
                var vse = stsd?._SampleEntry?.OfType<VisualSampleEntry>().FirstOrDefault();
                string codec = vse is not null ? FourCCToString(vse.FourCC) : "unknown";
                int width = vse?.Width ?? (int)((tkhd?.Width ?? 0) >> 16);
                int height = vse?.Height ?? (int)((tkhd?.Height ?? 0) >> 16);

                results.Add(new VideoTrackInfo(
                    TrackType: "video",
                    Codec: codec,
                    Width: width > 0 ? width : null,
                    Height: height > 0 ? height : null,
                    DurationSeconds: durationSeconds,
                    Bitrate: bitrate,
                    ChannelCount: null,
                    SampleRate: null));
            }
            else
            {
                var ase = stsd?._SampleEntry?.OfType<AudioSampleEntry>().FirstOrDefault();
                string codec = ase is not null ? FourCCToString(ase.FourCC) : "unknown";

                results.Add(new VideoTrackInfo(
                    TrackType: "audio",
                    Codec: codec,
                    Width: null,
                    Height: null,
                    DurationSeconds: durationSeconds,
                    Bitrate: bitrate,
                    ChannelCount: ase?.Channelcount > 0 ? (int)ase.Channelcount : null,
                    SampleRate: ase?.Samplerate > 0 ? (int)(ase.Samplerate >> 16) : null));
            }
        }

        return results;
    }

    private static double? ComputeDuration(
        MediaHeaderBox? mdhd,
        TrackHeaderBox? tkhd,
        uint movieTimescale,
        ulong movieDuration)
    {
        // Prefer media header (mdhd) timescale + duration.
        if (mdhd is not null && mdhd.Timescale > 0 && mdhd.Duration > 0)
            return (double)mdhd.Duration / mdhd.Timescale;

        // Fall back to track header duration + movie timescale.
        if (tkhd is not null && movieTimescale > 0 && tkhd.Duration > 0)
            return (double)tkhd.Duration / movieTimescale;

        // Last resort: movie-level duration.
        if (movieTimescale > 0 && movieDuration > 0)
            return (double)movieDuration / movieTimescale;

        return null;
    }

    private static bool IsHandlerType(uint fourCC, string expected)
    {
        return FourCCToString(fourCC) == expected;
    }

    private static string FourCCToString(uint fourCC)
    {
        var bytes = BitConverter.GetBytes(fourCC);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        return System.Text.Encoding.ASCII.GetString(bytes).TrimEnd('\0');
    }

    private static T? FindDescendant<T>(Box? box) where T : Box
    {
        if (box is null) return null;
        foreach (var child in box.Children)
        {
            if (child is T match) return match;
            var found = FindDescendant<T>(child);
            if (found is not null) return found;
        }
        return null;
    }
}
