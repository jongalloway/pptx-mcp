using System.CommandLine;
using PptxTools.Services;

namespace PptxTools.Commands;

/// <summary>CLI command group for managing slides.</summary>
public static class SlidesCommand
{
    public static Command Create(PresentationService service)
    {
        var command = new Command("slides") { Description = "Manage slides" };
        command.Add(CreateAddCommand(service));
        command.Add(CreateDeleteCommand(service));
        command.Add(CreateReorderCommand(service));
        command.Add(CreateDuplicateCommand(service));
        return command;
    }

    private static Command CreateAddCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var layoutOption = new Option<string?>("--layout") { Description = "Layout name to use for the new slide", DefaultValueFactory = _ => null };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("add") { Description = "Add a new slide to the presentation" };
        cmd.Add(fileArg);
        cmd.Add(layoutOption);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var layoutName = parseResult.GetValue(layoutOption);
            var asJson = parseResult.GetValue(jsonOption);

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                Environment.ExitCode = 1;
                return;
            }

            if (layoutName is not null)
            {
                var result = service.AddSlideFromLayout(filePath, layoutName);
                if (asJson)
                {
                    Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                    return;
                }
                Console.WriteLine(result.Message);
            }
            else
            {
                var slideNumber = service.AddSlide(filePath, null);
                if (asJson)
                {
                    Console.WriteLine(JsonSerializer.Serialize(new { SlideNumber = slideNumber }, JsonOptions));
                    return;
                }
                Console.WriteLine($"Added slide {slideNumber}");
            }
        });

        return cmd;
    }

    private static Command CreateDeleteCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var slideOption = new Option<int>("--slide") { Description = "Slide number to delete (1-based)", Required = true };

        var cmd = new Command("delete") { Description = "Delete a slide from the presentation" };
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

            service.DeleteSlide(filePath, slideNumber);
            Console.WriteLine($"Deleted slide {slideNumber}");
        });

        return cmd;
    }

    private static Command CreateReorderCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var orderOption = new Option<string>("--order") { Description = "New slide order as comma-separated numbers (e.g. 3,1,2)", Required = true };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("reorder") { Description = "Reorder slides in the presentation" };
        cmd.Add(fileArg);
        cmd.Add(orderOption);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var orderStr = parseResult.GetValue(orderOption)!;
            var asJson = parseResult.GetValue(jsonOption);

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                Environment.ExitCode = 1;
                return;
            }

            int[] newOrder;
            try
            {
                newOrder = orderStr.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Select(int.Parse)
                    .ToArray();
            }
            catch (FormatException)
            {
                Console.Error.WriteLine("Error: Invalid order format. Use comma-separated numbers (e.g. 3,1,2)");
                Environment.ExitCode = 1;
                return;
            }

            service.ReorderSlides(filePath, newOrder);

            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(new { Success = true, NewOrder = newOrder }, JsonOptions));
                return;
            }

            Console.WriteLine($"Reordered slides: {string.Join(", ", newOrder)}");
        });

        return cmd;
    }

    private static Command CreateDuplicateCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var slideOption = new Option<int>("--slide") { Description = "Slide number to duplicate (1-based)", Required = true };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("duplicate") { Description = "Duplicate a slide" };
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

            var result = service.DuplicateSlide(filePath, slideNumber);

            if (asJson)
            {
                Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                return;
            }

            Console.WriteLine(result.Message);
            if (result.Success)
            {
                Console.WriteLine($"  New slide number: {result.NewSlideNumber}");
                Console.WriteLine($"  Shapes copied: {result.ShapesCopied}");
            }
        });

        return cmd;
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
}
