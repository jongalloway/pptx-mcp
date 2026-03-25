using System.CommandLine;
using PptxTools.Models;
using PptxTools.Services;

namespace PptxTools.Commands;

/// <summary>CLI command group for editing presentation content.</summary>
public static class EditCommand
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly JsonSerializerOptions JsonReadOptions = new() { PropertyNameCaseInsensitive = true };

    public static Command Create(PresentationService service)
    {
        var command = new Command("edit") { Description = "Edit presentation content" };
        command.Add(CreateSlideCommand(service));
        command.Add(CreateBatchCommand(service));
        command.Add(CreateTableCommand(service));
        command.Add(CreateImageCommand(service));
        command.Add(CreateNotesCommand(service));
        command.Add(CreateChartCommand(service));
        return command;
    }

    // --- edit slide ---

    private static Command CreateSlideCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var slideOption = new Option<int>("--slide") { Description = "Slide number (1-based)", Required = true };
        var shapeOption = new Option<string>("--shape") { Description = "Shape name to update", Required = true };
        var textOption = new Option<string>("--text") { Description = "New text value", Required = true };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("slide") { Description = "Update text of a named shape on a slide" };
        cmd.Add(fileArg);
        cmd.Add(slideOption);
        cmd.Add(shapeOption);
        cmd.Add(textOption);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var slideNumber = parseResult.GetValue(slideOption);
            var shapeName = parseResult.GetValue(shapeOption)!;
            var text = parseResult.GetValue(textOption)!;
            var asJson = parseResult.GetValue(jsonOption);

            if (!ValidateFileExists(filePath)) return;

            try
            {
                var result = service.UpdateSlideData(filePath, slideNumber, shapeName, null, text);

                if (asJson)
                {
                    Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                    return;
                }

                if (result.Success)
                    Console.WriteLine($"Updated shape '{shapeName}' on slide {slideNumber}");
                else
                    WriteError(result.Message ?? "Update failed");
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
        });

        return cmd;
    }

    // --- edit batch ---

    private static Command CreateBatchCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var mutationsArg = new Argument<string>("mutations-json") { Description = "Path to JSON file containing mutations" };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("batch") { Description = "Apply batch text mutations from a JSON file" };
        cmd.Add(fileArg);
        cmd.Add(mutationsArg);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var mutationsPath = parseResult.GetValue(mutationsArg)!;
            var asJson = parseResult.GetValue(jsonOption);

            if (!ValidateFileExists(filePath)) return;

            if (!File.Exists(mutationsPath))
            {
                WriteError($"Mutations file not found: {mutationsPath}");
                return;
            }

            List<BatchUpdateMutation> mutations;
            try
            {
                var json = File.ReadAllText(mutationsPath);
                mutations = JsonSerializer.Deserialize<List<BatchUpdateMutation>>(json, JsonReadOptions)
                    ?? throw new InvalidOperationException("Deserialized mutations list was null");
            }
            catch (JsonException ex)
            {
                WriteError($"Invalid JSON in mutations file: {ex.Message}");
                return;
            }

            try
            {
                var result = service.BatchUpdate(filePath, mutations);

                if (asJson)
                {
                    Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                    return;
                }

                Console.WriteLine($"Batch update: {result.SuccessCount}/{result.TotalMutations} mutations applied");
                foreach (var mutation in result.Results)
                {
                    var status = mutation.Success ? "OK" : "FAIL";
                    Console.WriteLine($"  [{status}] Slide {mutation.SlideNumber}, shape '{mutation.ShapeName}'");
                    if (!mutation.Success && mutation.Error is not null)
                        Console.WriteLine($"         {mutation.Error}");
                }
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
        });

        return cmd;
    }

    // --- edit table (insert | update) ---

    private static Command CreateTableCommand(PresentationService service)
    {
        var tableCmd = new Command("table") { Description = "Insert or update tables" };
        tableCmd.Add(CreateTableInsertCommand(service));
        tableCmd.Add(CreateTableUpdateCommand(service));
        return tableCmd;
    }

    private static Command CreateTableInsertCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var slideOption = new Option<int>("--slide") { Description = "Slide number (1-based)", Required = true };
        var headersOption = new Option<string>("--headers") { Description = "Comma-separated column headers (e.g. \"Name,Value,Status\")", Required = true };
        var rowsOption = new Option<string>("--rows") { Description = "JSON array of row arrays (e.g. '[[\"Alice\",\"100\"],[\"Bob\",\"200\"]]')", Required = true };
        var nameOption = new Option<string?>("--name") { Description = "Optional table name", DefaultValueFactory = _ => null };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("insert") { Description = "Insert a new table on a slide" };
        cmd.Add(fileArg);
        cmd.Add(slideOption);
        cmd.Add(headersOption);
        cmd.Add(rowsOption);
        cmd.Add(nameOption);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var slideNumber = parseResult.GetValue(slideOption);
            var headersStr = parseResult.GetValue(headersOption)!;
            var rowsJson = parseResult.GetValue(rowsOption)!;
            var tableName = parseResult.GetValue(nameOption);
            var asJson = parseResult.GetValue(jsonOption);

            if (!ValidateFileExists(filePath)) return;

            var headers = headersStr.Split(',', StringSplitOptions.TrimEntries);

            string[][] rows;
            try
            {
                rows = JsonSerializer.Deserialize<string[][]>(rowsJson, JsonReadOptions)
                    ?? throw new InvalidOperationException("Deserialized rows array was null");
            }
            catch (JsonException ex)
            {
                WriteError($"Invalid JSON for --rows: {ex.Message}");
                return;
            }

            try
            {
                var result = service.InsertTable(filePath, slideNumber, headers, rows, tableName);

                if (asJson)
                {
                    Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                    return;
                }

                Console.WriteLine($"Inserted table on slide {slideNumber}: {headers.Length} columns, {rows.Length} rows");
                if (tableName is not null)
                    Console.WriteLine($"  Table name: {tableName}");
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
        });

        return cmd;
    }

    private static Command CreateTableUpdateCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var slideOption = new Option<int>("--slide") { Description = "Slide number (1-based)", Required = true };
        var updatesOption = new Option<string>("--updates") { Description = "JSON array of cell updates (e.g. '[{\"row\":0,\"column\":1,\"value\":\"New\"}]')", Required = true };
        var nameOption = new Option<string?>("--name") { Description = "Table name (case-insensitive)", DefaultValueFactory = _ => null };
        var indexOption = new Option<int?>("--index") { Description = "Zero-based table index on the slide", DefaultValueFactory = _ => null };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("update") { Description = "Update cells in an existing table" };
        cmd.Add(fileArg);
        cmd.Add(slideOption);
        cmd.Add(updatesOption);
        cmd.Add(nameOption);
        cmd.Add(indexOption);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var slideNumber = parseResult.GetValue(slideOption);
            var updatesJson = parseResult.GetValue(updatesOption)!;
            var tableName = parseResult.GetValue(nameOption);
            var tableIndex = parseResult.GetValue(indexOption);
            var asJson = parseResult.GetValue(jsonOption);

            if (!ValidateFileExists(filePath)) return;

            TableCellUpdate[] updates;
            try
            {
                updates = JsonSerializer.Deserialize<TableCellUpdate[]>(updatesJson, JsonReadOptions)
                    ?? throw new InvalidOperationException("Deserialized updates array was null");
            }
            catch (JsonException ex)
            {
                WriteError($"Invalid JSON for --updates: {ex.Message}");
                return;
            }

            try
            {
                var result = service.UpdateTable(filePath, slideNumber, updates, tableName, tableIndex);

                if (asJson)
                {
                    Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                    return;
                }

                Console.WriteLine($"Updated {updates.Length} cell(s) in table on slide {slideNumber}");
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
        });

        return cmd;
    }

    // --- edit image (insert | replace) ---

    private static Command CreateImageCommand(PresentationService service)
    {
        var imageCmd = new Command("image") { Description = "Insert or replace images" };
        imageCmd.Add(CreateImageInsertCommand(service));
        imageCmd.Add(CreateImageReplaceCommand(service));
        return imageCmd;
    }

    private static Command CreateImageInsertCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var slideOption = new Option<int>("--slide") { Description = "Slide number (1-based)", Required = true };
        var imageOption = new Option<string>("--image") { Description = "Path to the image file", Required = true };
        var xOption = new Option<long>("--x") { Description = "X position in EMUs", DefaultValueFactory = _ => 0 };
        var yOption = new Option<long>("--y") { Description = "Y position in EMUs", DefaultValueFactory = _ => 0 };
        var widthOption = new Option<long>("--width") { Description = "Width in EMUs", DefaultValueFactory = _ => 2743200 };
        var heightOption = new Option<long>("--height") { Description = "Height in EMUs", DefaultValueFactory = _ => 2057400 };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("insert") { Description = "Insert an image on a slide" };
        cmd.Add(fileArg);
        cmd.Add(slideOption);
        cmd.Add(imageOption);
        cmd.Add(xOption);
        cmd.Add(yOption);
        cmd.Add(widthOption);
        cmd.Add(heightOption);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var slideNumber = parseResult.GetValue(slideOption);
            var imagePath = parseResult.GetValue(imageOption)!;
            var x = parseResult.GetValue(xOption);
            var y = parseResult.GetValue(yOption);
            var width = parseResult.GetValue(widthOption);
            var height = parseResult.GetValue(heightOption);
            var asJson = parseResult.GetValue(jsonOption);

            if (!ValidateFileExists(filePath)) return;

            if (!File.Exists(imagePath))
            {
                WriteError($"Image file not found: {imagePath}");
                return;
            }

            try
            {
                // InsertImage uses 0-based slideIndex
                service.InsertImage(filePath, slideNumber - 1, imagePath, x, y, width, height);

                if (asJson)
                {
                    Console.WriteLine(JsonSerializer.Serialize(new { Success = true, SlideNumber = slideNumber, ImagePath = imagePath }, JsonOptions));
                    return;
                }

                Console.WriteLine($"Inserted image on slide {slideNumber}: {imagePath}");
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
        });

        return cmd;
    }

    private static Command CreateImageReplaceCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var slideOption = new Option<int>("--slide") { Description = "Slide number (1-based)", Required = true };
        var shapeOption = new Option<string>("--shape") { Description = "Picture shape name to replace", Required = true };
        var imageOption = new Option<string>("--image") { Description = "Path to the replacement image file", Required = true };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("replace") { Description = "Replace an existing image in a shape" };
        cmd.Add(fileArg);
        cmd.Add(slideOption);
        cmd.Add(shapeOption);
        cmd.Add(imageOption);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var slideNumber = parseResult.GetValue(slideOption);
            var shapeName = parseResult.GetValue(shapeOption)!;
            var imagePath = parseResult.GetValue(imageOption)!;
            var asJson = parseResult.GetValue(jsonOption);

            if (!ValidateFileExists(filePath)) return;

            if (!File.Exists(imagePath))
            {
                WriteError($"Image file not found: {imagePath}");
                return;
            }

            try
            {
                var result = service.ReplaceImage(filePath, slideNumber, shapeName, null, imagePath, null);

                if (asJson)
                {
                    Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                    return;
                }

                if (result.Success)
                    Console.WriteLine($"Replaced image in shape '{shapeName}' on slide {slideNumber}");
                else
                    WriteError(result.Message ?? "Image replacement failed");
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
        });

        return cmd;
    }

    // --- edit notes ---

    private static Command CreateNotesCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var slideOption = new Option<int>("--slide") { Description = "Slide number (1-based)", Required = true };
        var textOption = new Option<string>("--text") { Description = "Speaker notes text", Required = true };
        var appendOption = new Option<bool>("--append") { Description = "Append to existing notes instead of replacing" };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("notes") { Description = "Set or append speaker notes on a slide" };
        cmd.Add(fileArg);
        cmd.Add(slideOption);
        cmd.Add(textOption);
        cmd.Add(appendOption);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var slideNumber = parseResult.GetValue(slideOption);
            var text = parseResult.GetValue(textOption)!;
            var append = parseResult.GetValue(appendOption);
            var asJson = parseResult.GetValue(jsonOption);

            if (!ValidateFileExists(filePath)) return;

            try
            {
                // WriteNotes uses 0-based slideIndex
                service.WriteNotes(filePath, slideNumber - 1, text, append);

                if (asJson)
                {
                    Console.WriteLine(JsonSerializer.Serialize(new { Success = true, SlideNumber = slideNumber, Append = append }, JsonOptions));
                    return;
                }

                var action = append ? "Appended notes to" : "Set notes on";
                Console.WriteLine($"{action} slide {slideNumber}");
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
        });

        return cmd;
    }

    // --- edit chart ---

    private static Command CreateChartCommand(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var slideOption = new Option<int>("--slide") { Description = "Slide number (1-based)", Required = true };
        var dataOption = new Option<string>("--data") { Description = "JSON array of chart series updates", Required = true };
        var nameOption = new Option<string?>("--name") { Description = "Chart name (case-insensitive)", DefaultValueFactory = _ => null };
        var indexOption = new Option<int?>("--index") { Description = "Zero-based chart index on the slide", DefaultValueFactory = _ => null };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var cmd = new Command("chart") { Description = "Update chart data on a slide" };
        cmd.Add(fileArg);
        cmd.Add(slideOption);
        cmd.Add(dataOption);
        cmd.Add(nameOption);
        cmd.Add(indexOption);
        cmd.Add(jsonOption);

        cmd.SetAction(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var slideNumber = parseResult.GetValue(slideOption);
            var dataJson = parseResult.GetValue(dataOption)!;
            var chartName = parseResult.GetValue(nameOption);
            var chartIndex = parseResult.GetValue(indexOption);
            var asJson = parseResult.GetValue(jsonOption);

            if (!ValidateFileExists(filePath)) return;

            ChartSeriesUpdate[] updates;
            try
            {
                updates = JsonSerializer.Deserialize<ChartSeriesUpdate[]>(dataJson, JsonReadOptions)
                    ?? throw new InvalidOperationException("Deserialized chart data array was null");
            }
            catch (JsonException ex)
            {
                WriteError($"Invalid JSON for --data: {ex.Message}");
                return;
            }

            try
            {
                var result = service.UpdateChartData(filePath, slideNumber, updates, chartName, chartIndex);

                if (asJson)
                {
                    Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                    return;
                }

                if (result.Success)
                    Console.WriteLine($"Updated chart data on slide {slideNumber}: {updates.Length} series update(s)");
                else
                    WriteError(result.Message ?? "Chart update failed");
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
        });

        return cmd;
    }

    // --- Helpers ---

    private static bool ValidateFileExists(string filePath)
    {
        if (File.Exists(filePath)) return true;
        WriteError($"File not found: {filePath}");
        return false;
    }

    private static void WriteError(string message)
    {
        Console.Error.WriteLine($"Error: {message}");
        Environment.ExitCode = 1;
    }
}
