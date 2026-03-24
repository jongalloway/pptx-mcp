using System.CommandLine;
using PptxMcp.Services;

namespace PptxMcp.Commands;

/// <summary>CLI command group for managing media assets.</summary>
public static class MediaCommand
{
    public static Command Create(PresentationService service)
    {
        var command = new Command("media") { Description = "Manage media assets" };
        command.Add(CreateAnalyzeCommand(service));
        command.Add(CreateDeduplicateCommand(service));
        command.Add(CreateAnalyzeVideoCommand(service));
        return command;
    }

    private static Command CreateAnalyzeCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("analyze") { Description = "Analyze media inventory in the presentation" };
        cmd.Add(fileArg);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var asJson = parseResult.GetValue(jsonOption);

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                Environment.ExitCode = 1;
                return;
            }

            var result = service.AnalyzeMedia(filePath);

            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                return;
            }

            Console.WriteLine($"Media Analysis: {filePath}");
            Console.WriteLine($"  Total media: {result.TotalMediaCount} parts, {FormatBytes(result.TotalMediaSize)}");
            Console.WriteLine($"  Duplicate groups: {result.DuplicateGroupCount} (potential savings: {FormatBytes(result.DuplicateSavingsBytes)})");
            Console.WriteLine();

            if (result.MediaParts.Length > 0)
            {
                Console.WriteLine("Media parts:");
                foreach (var part in result.MediaParts)
                {
                    var slides = part.ReferencedBySlides.Length > 0
                        ? $"slides {string.Join(", ", part.ReferencedBySlides)}"
                        : "no slide references";
                    Console.WriteLine($"  {part.Path}  {FormatBytes(part.SizeBytes)}  [{part.ContentType}]  ({slides})");
                }
            }

            if (result.DuplicateGroups.Length > 0)
            {
                Console.WriteLine();
                Console.WriteLine("Duplicate groups:");
                foreach (var group in result.DuplicateGroups)
                {
                    Console.WriteLine($"  [{group.ContentType}] {FormatBytes(group.SizeBytes)} x {group.Parts.Length} copies");
                    foreach (var part in group.Parts)
                        Console.WriteLine($"    {part}");
                }
            }
        });

        return cmd;
    }

    private static Command CreateDeduplicateCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("deduplicate") { Description = "Remove duplicate media from the presentation" };
        cmd.Add(fileArg);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var asJson = parseResult.GetValue(jsonOption);

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                Environment.ExitCode = 1;
                return;
            }

            var result = service.DeduplicateMedia(filePath);

            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                return;
            }

            Console.WriteLine(result.Message);
            if (result.Success && result.PartsRemoved > 0)
            {
                Console.WriteLine($"  Groups found: {result.DuplicateGroupsFound}");
                Console.WriteLine($"  Parts removed: {result.PartsRemoved}");
                Console.WriteLine($"  Bytes saved: {FormatBytes(result.BytesSaved)}");
                Console.WriteLine($"  Validation: {(result.Validation.IsValid ? "passed" : "failed")}");
            }
        });

        return cmd;
    }

    private static Command CreateAnalyzeVideoCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("analyze-video") { Description = "Analyze video and audio metadata" };
        cmd.Add(fileArg);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var asJson = parseResult.GetValue(jsonOption);

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                Environment.ExitCode = 1;
                return;
            }

            var result = service.AnalyzeVideoMetadata(filePath);

            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                return;
            }

            Console.WriteLine($"Video/Audio Analysis: {filePath}");
            Console.WriteLine($"  Parts found: {result.VideoPartsFound}");
            Console.WriteLine($"  Total tracks: {result.TotalTracks}");
            Console.WriteLine();

            foreach (var part in result.Parts)
            {
                Console.WriteLine($"  {part.PartUri}  [{part.ContentType}]  {FormatBytes(part.FileSizeBytes)}");
                if (part.Error is not null)
                {
                    Console.WriteLine($"    Error: {part.Error}");
                    continue;
                }
                foreach (var track in part.Tracks)
                {
                    var details = track.TrackType switch
                    {
                        "Video" => $"{track.Codec} {track.Width}x{track.Height}" +
                                   (track.DurationSeconds.HasValue ? $" {track.DurationSeconds:F1}s" : "") +
                                   (track.Bitrate.HasValue ? $" {FormatBytes(track.Bitrate.Value)}/s" : ""),
                        "Audio" => $"{track.Codec}" +
                                   (track.ChannelCount.HasValue ? $" {track.ChannelCount}ch" : "") +
                                   (track.SampleRate.HasValue ? $" {track.SampleRate}Hz" : "") +
                                   (track.DurationSeconds.HasValue ? $" {track.DurationSeconds:F1}s" : ""),
                        _ => track.Codec
                    };
                    Console.WriteLine($"    [{track.TrackType}] {details}");
                }
            }
        });

        return cmd;
    }

    private static string FormatBytes(long bytes)
    {
        return bytes switch
        {
            >= 1_073_741_824 => $"{bytes / 1_073_741_824.0:F1} GB",
            >= 1_048_576 => $"{bytes / 1_048_576.0:F1} MB",
            >= 1_024 => $"{bytes / 1_024.0:F1} KB",
            _ => $"{bytes} B"
        };
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
}
