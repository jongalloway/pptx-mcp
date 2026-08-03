using System.CommandLine;
using PptxTools.Services;

namespace PptxTools.Commands;

/// <summary>CLI command group for presentation export.</summary>
public static class ExportCommand
{
    public static Command Create(PresentationService service)
    {
        var command = new Command("export") { Description = "Export presentation content" };
        command.Add(CreateMarkdownCommand(service));
        return command;
    }

    private static Command CreateMarkdownCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var outputArg = new Argument<string?>("output") { Description = "Output file path (defaults to stdout)", DefaultValueFactory = _ => null };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("markdown") { Description = "Export presentation as Markdown" };
        cmd.Add(fileArg);
        cmd.Add(outputArg);
        cmd.Add(jsonOption);

        cmd.SetAction((Func<ParseResult, int>)(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var outputPath = parseResult.GetValue(outputArg);
            var asJson = parseResult.GetValue(jsonOption);

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                return 1;
            }

            var result = service.ExportMarkdown(filePath, outputPath);

            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                return 0;
            }

            if (outputPath is not null)
            {
                Console.WriteLine($"Exported {result.SlideCount} slides to {result.OutputPath}");
                if (result.ImageCount > 0)
                    Console.WriteLine($"Extracted {result.ImageCount} images");
            }
            else
            {
                Console.Write(result.Markdown);
            }

            return 0;
        }));

        return cmd;
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
}
