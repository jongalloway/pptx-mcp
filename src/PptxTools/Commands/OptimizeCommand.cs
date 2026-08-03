using System.CommandLine;
using System.Text.Json;
using PptxTools.Services;

namespace PptxTools.Commands;

/// <summary>CLI command for optimizing presentation file size.</summary>
public static class OptimizeCommand
{
    public static Command Create(PresentationService service)
    {
        var fileArg = new Argument<string>("file") { Description = "Path to the .pptx file" };
        var outputOption = new Option<string?>("--output") { Description = "Output file path (default: {name}_optimized.pptx)" };
        var qualityOption = new Option<int>("--quality") { Description = "Image compression quality 1-100", DefaultValueFactory = _ => 85 };
        var dpiOption = new Option<int>("--dpi") { Description = "Target DPI for image downscaling", DefaultValueFactory = _ => 150 };
        var removeLayoutsOption = new Option<bool>("--remove-layouts") { Description = "Remove unused layouts (default: true)", DefaultValueFactory = _ => true };
        var dedupMediaOption = new Option<bool>("--dedup-media") { Description = "Deduplicate media (default: true)", DefaultValueFactory = _ => true };
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON" };

        var command = new Command("optimize") { Description = "Optimize presentation file size" };
        command.Add(fileArg);
        command.Add(outputOption);
        command.Add(qualityOption);
        command.Add(dpiOption);
        command.Add(removeLayoutsOption);
        command.Add(dedupMediaOption);
        command.Add(jsonOption);

        // Support --no-remove-layouts and --no-dedup-media
        var noRemoveLayoutsOption = new Option<bool>("--no-remove-layouts") { Description = "Disable layout removal" };
        var noDedupMediaOption = new Option<bool>("--no-dedup-media") { Description = "Disable media deduplication" };
        command.Add(noRemoveLayoutsOption);
        command.Add(noDedupMediaOption);

        command.SetAction((Func<ParseResult, int>)(parseResult =>
        {
            var filePath = parseResult.GetValue(fileArg)!;
            var outputPath = parseResult.GetValue(outputOption);
            var quality = parseResult.GetValue(qualityOption);
            var dpi = parseResult.GetValue(dpiOption);
            var removeLayouts = parseResult.GetValue(removeLayoutsOption);
            var dedupMedia = parseResult.GetValue(dedupMediaOption);
            var asJson = parseResult.GetValue(jsonOption);

            // Handle --no-* overrides
            if (parseResult.GetValue(noRemoveLayoutsOption))
                removeLayouts = false;
            if (parseResult.GetValue(noDedupMediaOption))
                dedupMedia = false;

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                return 1;
            }

            // Default output path: {name}_optimized.pptx alongside original
            if (string.IsNullOrEmpty(outputPath))
            {
                var dir = Path.GetDirectoryName(filePath) ?? ".";
                var name = Path.GetFileNameWithoutExtension(filePath);
                var ext = Path.GetExtension(filePath);
                outputPath = Path.Join(dir, $"{name}_optimized{ext}");
            }

            return RunOptimize(service, filePath, outputPath, quality, dpi, removeLayouts, dedupMedia, asJson);
        }));

        return command;
    }

    private static int RunOptimize(
        PresentationService service,
        string filePath,
        string outputPath,
        int quality,
        int dpi,
        bool removeLayouts,
        bool dedupMedia,
        bool asJson)
    {
        var inputName = Path.GetFileName(filePath);
        var outputName = Path.GetFileName(outputPath);

        if (!asJson)
            Console.WriteLine($"Optimizing: {inputName} → {outputName}");

        // Step 1: Copy input to output (never modify original)
        try
        {
            File.Copy(filePath, outputPath, overwrite: true);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: Could not copy file: {ex.Message}");
            return 1;
        }

        // Step 2: Capture "before" size
        var beforeResult = service.AnalyzeFileSize(outputPath);
        long beforeSize = beforeResult.TotalFileSize;

        if (!asJson)
        {
            Console.WriteLine();
            Console.WriteLine($"Before: {FormatBytes(beforeSize)}");
            Console.WriteLine();
        }

        // Track step results for JSON output
        var steps = new List<OptimizeStepResult>();

        // Step 3: Deduplicate media
        if (dedupMedia)
        {
            try
            {
                var dedupResult = service.DeduplicateMedia(outputPath);
                var step = new OptimizeStepResult(
                    "dedup-media",
                    true,
                    dedupResult.PartsRemoved,
                    dedupResult.BytesSaved,
                    $"{dedupResult.PartsRemoved} duplicates removed",
                    null);
                steps.Add(step);

                if (!asJson)
                    Console.WriteLine($"  ✓ Deduplicated media: {dedupResult.PartsRemoved} duplicates removed (saved {FormatBytes(dedupResult.BytesSaved)})");
            }
            catch (Exception ex)
            {
                steps.Add(new OptimizeStepResult("dedup-media", false, 0, 0, null, ex.Message));
                if (!asJson)
                    Console.WriteLine($"  ✗ Deduplicate media failed: {ex.Message}");
            }
        }

        // Step 4: Optimize images
        try
        {
            var imageResult = service.OptimizeImages(outputPath, dpi, quality);
            var step = new OptimizeStepResult(
                "optimize-images",
                true,
                imageResult.ImagesProcessed,
                imageResult.TotalBytesSaved,
                $"{imageResult.ImagesProcessed} images compressed",
                null);
            steps.Add(step);

            if (!asJson)
                Console.WriteLine($"  ✓ Optimized images: {imageResult.ImagesProcessed} images compressed (saved {FormatBytes(imageResult.TotalBytesSaved)})");
        }
        catch (Exception ex)
        {
            steps.Add(new OptimizeStepResult("optimize-images", false, 0, 0, null, ex.Message));
            if (!asJson)
                Console.WriteLine($"  ✗ Optimize images failed: {ex.Message}");
        }

        // Step 5: Remove unused layouts
        if (removeLayouts)
        {
            try
            {
                var layoutResult = service.RemoveUnusedLayouts(outputPath, null);
                var step = new OptimizeStepResult(
                    "remove-layouts",
                    true,
                    layoutResult.LayoutsRemoved,
                    layoutResult.BytesSaved,
                    $"{layoutResult.LayoutsRemoved} unused layouts removed",
                    null);
                steps.Add(step);

                if (!asJson)
                    Console.WriteLine($"  ✓ Removed layouts: {layoutResult.LayoutsRemoved} unused layouts removed (saved {FormatBytes(layoutResult.BytesSaved)})");
            }
            catch (Exception ex)
            {
                steps.Add(new OptimizeStepResult("remove-layouts", false, 0, 0, null, ex.Message));
                if (!asJson)
                    Console.WriteLine($"  ✗ Remove layouts failed: {ex.Message}");
            }
        }

        // Step 6: Capture "after" size
        var afterResult = service.AnalyzeFileSize(outputPath);
        long afterSize = afterResult.TotalFileSize;
        long savings = beforeSize - afterSize;
        double savingsPercent = beforeSize > 0 ? (double)savings / beforeSize * 100 : 0;

        if (asJson)
        {
            var jsonResult = new OptimizeResult(
                filePath,
                outputPath,
                beforeSize,
                afterSize,
                savings,
                Math.Round(savingsPercent, 1),
                steps);

            Console.WriteLine(JsonSerializer.Serialize(jsonResult, JsonOptions));
        }
        else
        {
            Console.WriteLine();
            Console.WriteLine($"After: {FormatBytes(afterSize)}");
            Console.WriteLine($"Savings: {FormatBytes(savings)} ({savingsPercent:F1}%)");
        }

        return 0;
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

    // JSON output models
    private record OptimizeStepResult(
        string Step,
        bool Success,
        int ItemsAffected,
        long BytesSaved,
        string? Summary,
        string? Error);

    private record OptimizeResult(
        string InputFile,
        string OutputFile,
        long BeforeSize,
        long AfterSize,
        long Savings,
        double SavingsPercent,
        List<OptimizeStepResult> Steps);
}
