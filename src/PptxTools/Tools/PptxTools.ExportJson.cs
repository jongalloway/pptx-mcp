using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxTools.Models;

namespace PptxTools.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Export a PowerPoint presentation to structured JSON.
    /// Available actions:
    /// - Full: Export everything — metadata, all slides with shapes, tables, charts, images, and notes.
    /// - SlidesOnly: Export only slide content (shapes, tables, charts, images, notes) without metadata.
    /// - MetadataOnly: Export only presentation-level metadata (title, author, dates, etc.).
    /// - SchemaOnly: Return the JSON schema description without reading any file.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file. Not required for SchemaOnly action.</param>
    /// <param name="action">The export scope: Full, SlidesOnly, MetadataOnly, or SchemaOnly.</param>
    [McpServerTool(Title = "Export JSON", ReadOnly = true, Idempotent = true)]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["Full","SlidesOnly","MetadataOnly","SchemaOnly"]""")]
    public partial Task<string> pptx_export_json(
        string? filePath,
        ExportJsonAction action)
    {
        if (action == ExportJsonAction.SchemaOnly)
        {
            var schema = _service.ExportJson("", ExportJsonAction.SchemaOnly);
            return Task.FromResult(JsonSerializer.Serialize(schema, IndentedJson));
        }

        if (string.IsNullOrWhiteSpace(filePath))
            return Task.FromResult(JsonSerializer.Serialize(
                MakeExportJsonError(action, filePath, "filePath is required for this action."), IndentedJson));

        return ExecuteToolStructured(filePath!,
            () => _service.ExportJson(filePath!, action),
            error => MakeExportJsonError(action, filePath, error));
    }

    private static PresentationExport MakeExportJsonError(ExportJsonAction action, string? filePath, string message) =>
        new(
            Success: false,
            Action: action.ToString(),
            FilePath: filePath,
            Metadata: null,
            SlideCount: 0,
            Slides: null,
            Schema: null,
            Message: message);
}
