using System.CommandLine;
using PptxMcp.Services;

namespace PptxMcp.Commands;

/// <summary>CLI command group for presentation analysis.</summary>
public static class AnalyzeCommand
{
    public static Command Create(PresentationService service)
    {
        var command = new Command("analyze") { Description = "Analyze presentation structure and content" };
        command.Add(CreateFileSizeCommand(service));
        command.Add(CreateTalkingPointsCommand(service));
        return command;
    }

    private static Command CreateFileSizeCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("file-size") { Description = "Analyze file size breakdown by category" };
        cmd.Add(fileArg);
        cmd.Add(jsonOption);

        cmd.SetAction((Func<ParseResult, int>)(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var asJson = parseResult.GetValue(jsonOption);

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                return 1;
            }

            var result = service.AnalyzeFileSize(filePath);

            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                return 0;
            }

            Console.WriteLine($"File: {result.FilePath}");
            Console.WriteLine($"Total file size: {FormatBytes(result.TotalFileSize)}");
            Console.WriteLine($"Total part size: {FormatBytes(result.TotalPartSize)}");
            Console.WriteLine();

            foreach (var category in result.Categories)
            {
                if (category.PartCount == 0)
                    continue;

                Console.WriteLine($"  {category.Name} ({category.PartCount} parts, {FormatBytes(category.TotalSize)})");
                foreach (var part in category.Parts)
                {
                    Console.WriteLine($"    {part.Path}  {FormatBytes(part.Size)}  [{part.ContentType}]");
                }
            }

            return 0;
        }));

        return cmd;
    }

    private static Command CreateTalkingPointsCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var topOption = new Option<int>("--top") { Description = "Number of talking points per slide", DefaultValueFactory = _ => 5 };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("talking-points") { Description = "Extract key talking points from slides" };
        cmd.Add(fileArg);
        cmd.Add(topOption);
        cmd.Add(jsonOption);

        cmd.SetAction((Func<ParseResult, int>)(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var topN = parseResult.GetValue(topOption);
            var asJson = parseResult.GetValue(jsonOption);

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                return 1;
            }

            var result = service.ExtractTalkingPoints(filePath, topN);

            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                return 0;
            }

            foreach (var slide in result)
            {
                var title = slide.Title ?? "(untitled)";
                Console.WriteLine($"Slide {slide.SlideIndex + 1}: {title}");
                foreach (var point in slide.Points)
                {
                    Console.WriteLine($"  • {point}");
                }
                if (slide.Points.Count == 0)
                    Console.WriteLine("  (no talking points)");
                Console.WriteLine();
            }

            return 0;
        }));

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
