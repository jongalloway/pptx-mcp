using System.CommandLine;
using PptxTools.Services;

namespace PptxTools.Commands;

/// <summary>CLI command group for inspecting slide details and metadata.</summary>
public static class InspectCommand
{
    public static Command Create(PresentationService service)
    {
        var command = new Command("inspect") { Description = "Inspect slide details and metadata" };
        command.Add(CreateSlidesCommand(service));
        command.Add(CreateContentCommand(service));
        command.Add(CreateXmlCommand(service));
        command.Add(CreateLayoutsCommand(service));
        return command;
    }

    private static Command CreateSlidesCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("slides") { Description = "List all slides in the presentation" };
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

            var slides = service.GetSlides(filePath);

            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(slides, JsonOptions));
                return;
            }

            Console.WriteLine($"Slides ({slides.Count}):");
            Console.WriteLine();
            foreach (var slide in slides)
            {
                var title = slide.Title ?? "(untitled)";
                Console.WriteLine($"  Slide {slide.Index + 1}: {title}");
                Console.WriteLine($"    Placeholders: {slide.PlaceholderCount}");
                if (slide.Notes is not null)
                    Console.WriteLine($"    Notes: {Truncate(slide.Notes, 80)}");
            }
        });

        return cmd;
    }

    private static Command CreateContentCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var slideOption = new Option<int>("--slide") { Description = "Slide number (1-based)", Required = true };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("content") { Description = "Get detailed shape content for a slide" };
        cmd.Add(fileArg);
        cmd.Add(slideOption);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var slideNumber = parseResult.GetValue(slideOption);
            var asJson = parseResult.GetValue(jsonOption);

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                Environment.ExitCode = 1;
                return;
            }

            // Service uses 0-based slideIndex
            var content = service.GetSlideContent(filePath, slideNumber - 1);

            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(content, JsonOptions));
                return;
            }

            Console.WriteLine($"Slide {slideNumber} ({content.Shapes.Count} shapes)");
            Console.WriteLine($"  Size: {content.SlideWidthEmu} x {content.SlideHeightEmu} EMU");
            Console.WriteLine();

            foreach (var shape in content.Shapes)
            {
                Console.WriteLine($"  [{shape.ShapeType}] {shape.Name}");
                if (shape.IsPlaceholder)
                    Console.WriteLine($"    Placeholder: {shape.PlaceholderType ?? "unknown"} (index {shape.PlaceholderIndex})");
                if (shape.X is not null)
                    Console.WriteLine($"    Position: ({shape.X}, {shape.Y}) Size: ({shape.Width}, {shape.Height})");
                if (shape.Text is not null)
                    Console.WriteLine($"    Text: {Truncate(shape.Text, 120)}");
                if (shape.TableRows is not null)
                    Console.WriteLine($"    Table: {shape.TableRows.Count} rows");
            }
        });

        return cmd;
    }

    private static Command CreateXmlCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var slideOption = new Option<int>("--slide") { Description = "Slide number (1-based)", Required = true };

        var cmd = new Command("xml") { Description = "Get raw slide XML" };
        cmd.Add(fileArg);
        cmd.Add(slideOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var slideNumber = parseResult.GetValue(slideOption);

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                Environment.ExitCode = 1;
                return;
            }

            // Service uses 0-based slideIndex
            var xml = service.GetSlideXml(filePath, slideNumber - 1);
            Console.WriteLine(xml);
        });

        return cmd;
    }

    private static Command CreateLayoutsCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("layouts") { Description = "List all available slide layouts" };
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

            var layouts = service.GetLayouts(filePath);

            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(layouts, JsonOptions));
                return;
            }

            Console.WriteLine($"Layouts ({layouts.Count}):");
            Console.WriteLine();
            foreach (var layout in layouts)
            {
                Console.WriteLine($"  {layout.Index + 1}. {layout.Name}");
            }
        });

        return cmd;
    }

    private static string Truncate(string text, int maxLength)
    {
        var singleLine = text.ReplaceLineEndings(" ");
        return singleLine.Length <= maxLength ? singleLine : string.Concat(singleLine.AsSpan(0, maxLength - 3), "...");
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
}
