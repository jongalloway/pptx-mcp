using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxTools.Models;

namespace PptxTools.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Compare two PowerPoint presentations and identify differences.
    /// Available actions:
    /// - Full: Run all comparison checks (slides, text, metadata).
    /// - SlidesOnly: Compare only slide-level changes (added, removed).
    /// - TextOnly: Compare only text content changes across matching slides.
    /// - MetadataOnly: Compare only presentation-level metadata fields.
    /// </summary>
    /// <param name="originalFilePath">Absolute or relative path to the source .pptx file.</param>
    /// <param name="modifiedFilePath">Absolute or relative path to the target .pptx file.</param>
    /// <param name="action">The comparison scope: Full, SlidesOnly, TextOnly, or MetadataOnly.</param>
    [McpServerTool(Title = "Compare Presentations", ReadOnly = true, Idempotent = true)]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["Full","SlidesOnly","TextOnly","MetadataOnly"]""")]
    public partial Task<string> pptx_compare_presentations(
        string originalFilePath,
        string modifiedFilePath,
        CompareAction action)
    {
        ComparisonResult makeError(string message) => new(
            Success: false,
            Action: action.ToString(),
            SourceFile: originalFilePath,
            TargetFile: modifiedFilePath,
            AreIdentical: false,
            DifferenceCount: 0,
            SlideDifferences: null,
            TextDifferences: null,
            MetadataDifferences: null,
            Message: message);

        if (!File.Exists(originalFilePath))
            return Task.FromResult(JsonSerializer.Serialize(makeError($"File not found: {originalFilePath}"), IndentedJson));
        if (!File.Exists(modifiedFilePath))
            return Task.FromResult(JsonSerializer.Serialize(makeError($"File not found: {modifiedFilePath}"), IndentedJson));

        try
        {
            var result = _service.ComparePresentations(originalFilePath, modifiedFilePath, action);
            return Task.FromResult(JsonSerializer.Serialize(result, IndentedJson));
        }
        catch (Exception ex)
        {
            return Task.FromResult(JsonSerializer.Serialize(makeError($"Error: {ex.Message}"), IndentedJson));
        }
    }
}
