using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using PptxTools.Services;

namespace PptxTools.Completions;

/// <summary>
/// Provides argument auto-completions for pptx-tools resource templates and prompts.
/// Registered via <c>.WithCompleteHandler(...)</c> in Program.cs.
/// </summary>
public static class PptxCompletionHandler
{
    /// <summary>Known placeholder type values from the OpenXML PlaceholderValues enum.</summary>
    private static readonly string[] KnownPlaceholderTypes =
    [
        "title", "body", "ctrTitle", "subTitle", "dt", "ftr", "sldNum",
        "hdr", "obj", "pic", "tbl", "chart", "dgm", "media", "clipArt"
    ];

    /// <summary>Common action enum values used across tools.</summary>
    private static readonly string[] KnownActions =
    [
        "Add", "AddFromLayout", "Duplicate", "Move", "Reorder", "Find",
        "Remove", "Analyze", "Deduplicate", "AnalyzeVideo", "Read", "Update"
    ];

    /// <summary>Image/export format options accepted by insert/replace/export tools.</summary>
    private static readonly string[] KnownFormats =
    [
        "png", "jpg", "jpeg", "gif", "bmp", "svg", "markdown"
    ];

    /// <summary>Speaker notes style options used by prompts.</summary>
    private static readonly string[] KnownStyles =
    [
        "bullet-points", "narrative", "timing-cues"
    ];

    /// <summary>Chart-specific action values.</summary>
    private static readonly string[] KnownChartActions =
    [
        "Read", "Update"
    ];

    /// <summary>
    /// Handles <c>completion/complete</c> requests by returning matching values for
    /// layout names, shape names, and placeholder types.
    /// </summary>
    public static ValueTask<CompleteResult> HandleAsync(
        RequestContext<CompleteRequestParams> context,
        CancellationToken cancellationToken)
    {
        var request = context.Params;
        if (request is null)
            return ValueTask.FromResult(GetCompletions(null, string.Empty, null, null));

        var service = context.Services?.GetService(typeof(PresentationService)) as PresentationService;

        return ValueTask.FromResult(GetCompletions(
            request.Argument?.Name,
            request.Argument?.Value ?? string.Empty,
            request.Context?.Arguments,
            service));
    }

    /// <summary>
    /// Core completion logic — separated for testability.
    /// </summary>
    /// <param name="argumentName">Name of the argument being completed.</param>
    /// <param name="partialValue">Current partial value entered by the user.</param>
    /// <param name="contextArgs">
    /// Previously resolved template/prompt arguments. May contain a "file" key (from resource
    /// template variables) or a "filePath" key (from prompt arguments) with the path to the .pptx file.
    /// </param>
    /// <param name="service">Optional <see cref="PresentationService"/> for live data lookups.</param>
    public static CompleteResult GetCompletions(
        string? argumentName,
        string partialValue,
        IDictionary<string, string>? contextArgs,
        PresentationService? service)
    {
        if (argumentName is null)
            return EmptyResult();

        // --- Static completions (no file context needed) ---

        if (argumentName.Equals("placeholderType", StringComparison.OrdinalIgnoreCase))
            return FilterCompletions(KnownPlaceholderTypes, partialValue);

        if (argumentName.Equals("action", StringComparison.OrdinalIgnoreCase))
            return FilterCompletions(KnownActions, partialValue);

        if (argumentName.Equals("format", StringComparison.OrdinalIgnoreCase))
            return FilterCompletions(KnownFormats, partialValue);

        if (argumentName.Equals("style", StringComparison.OrdinalIgnoreCase))
            return FilterCompletions(KnownStyles, partialValue);

        if (argumentName.Equals("chartAction", StringComparison.OrdinalIgnoreCase))
            return FilterCompletions(KnownChartActions, partialValue);

        // --- Dynamic completions (require a file path) ---
        // The file path can come from:
        //   - The "file" argument on resource templates
        //   - The "filePath" argument on prompts
        // We inspect contextArgs for a previously resolved file argument.
        string? resolvedFilePath = null;
        if (contextArgs is not null)
        {
            contextArgs.TryGetValue("file", out resolvedFilePath);
            if (string.IsNullOrWhiteSpace(resolvedFilePath))
                contextArgs.TryGetValue("filePath", out resolvedFilePath);
            if (!string.IsNullOrWhiteSpace(resolvedFilePath))
                resolvedFilePath = Uri.UnescapeDataString(resolvedFilePath);
        }

        if (argumentName.Equals("layoutName", StringComparison.OrdinalIgnoreCase)
            || argumentName.Equals("layout", StringComparison.OrdinalIgnoreCase))
        {
            if (service is null || string.IsNullOrWhiteSpace(resolvedFilePath) || !File.Exists(resolvedFilePath))
                return EmptyResult();

            try
            {
                var layouts = service.GetLayouts(resolvedFilePath);
                var names = layouts.Select(l => l.Name).ToArray();
                return FilterCompletions(names, partialValue);
            }
            catch
            {
                return EmptyResult();
            }
        }

        if (argumentName.Equals("shapeName", StringComparison.OrdinalIgnoreCase)
            || argumentName.Equals("shape", StringComparison.OrdinalIgnoreCase))
        {
            if (service is null || string.IsNullOrWhiteSpace(resolvedFilePath) || !File.Exists(resolvedFilePath))
                return EmptyResult();

            try
            {
                var allSlides = service.GetAllSlideContents(resolvedFilePath);
                var uniqueNames = allSlides
                    .SelectMany(s => s.Shapes.Select(sh => sh.Name))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                return FilterCompletions(uniqueNames, partialValue);
            }
            catch
            {
                return EmptyResult();
            }
        }

        if (argumentName.Equals("slideNumber", StringComparison.OrdinalIgnoreCase)
            || argumentName.Equals("slideIndex", StringComparison.OrdinalIgnoreCase))
        {
            if (service is null || string.IsNullOrWhiteSpace(resolvedFilePath) || !File.Exists(resolvedFilePath))
                return EmptyResult();

            try
            {
                var slides = service.GetSlides(resolvedFilePath);
                var numbers = Enumerable.Range(1, slides.Count).Select(n => n.ToString()).ToArray();
                return FilterCompletions(numbers, partialValue);
            }
            catch
            {
                return EmptyResult();
            }
        }

        if (argumentName.Equals("tableName", StringComparison.OrdinalIgnoreCase)
            || argumentName.Equals("table", StringComparison.OrdinalIgnoreCase))
        {
            if (service is null || string.IsNullOrWhiteSpace(resolvedFilePath) || !File.Exists(resolvedFilePath))
                return EmptyResult();

            try
            {
                var allSlides = service.GetAllSlideContents(resolvedFilePath);
                var uniqueNames = allSlides
                    .SelectMany(s => s.Shapes
                        .Where(sh => sh.ShapeType.Equals("Table", StringComparison.OrdinalIgnoreCase))
                        .Select(sh => sh.Name))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                return FilterCompletions(uniqueNames, partialValue);
            }
            catch
            {
                return EmptyResult();
            }
        }

        return EmptyResult();
    }

    private static CompleteResult FilterCompletions(string[] candidates, string partialValue)
    {
        var matches = string.IsNullOrEmpty(partialValue)
            ? candidates
            : candidates.Where(c => c.StartsWith(partialValue, StringComparison.OrdinalIgnoreCase)).ToArray();

        return new CompleteResult
        {
            Completion = new Completion
            {
                Values = [.. matches],
                Total = matches.Length,
                HasMore = false
            }
        };
    }

    private static CompleteResult EmptyResult() =>
        new()
        {
            Completion = new Completion
            {
                Values = [],
                Total = 0,
                HasMore = false
            }
        };
}
