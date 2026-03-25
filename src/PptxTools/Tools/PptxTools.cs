using System.Text.Json;
using ModelContextProtocol.Server;
using PptxTools.Models;
using PptxTools.Services;

namespace PptxTools.Tools;

[McpServerToolType]
public sealed partial class PptxTools
{
    private static readonly JsonSerializerOptions IndentedJson = new() { WriteIndented = true };

    private readonly PresentationService _service;

    public PptxTools(PresentationService service)
    {
        _service = service;
    }

    /// <summary>File-check + try-catch wrapper for tools that return a plain string.</summary>
    private static Task<string> ExecuteTool(string filePath, Func<string> action)
    {
        if (!File.Exists(filePath))
            return Task.FromResult($"Error: File not found: {filePath}");
        try
        {
            return Task.FromResult(action());
        }
        catch (Exception ex)
        {
            return Task.FromResult($"Error: {ex.Message}");
        }
    }

    /// <summary>File-check + try-catch + JSON serialization for tools that return a typed result.</summary>
    private static Task<string> ExecuteToolJson<T>(string filePath, Func<T> action)
    {
        if (!File.Exists(filePath))
            return Task.FromResult($"Error: File not found: {filePath}");
        try
        {
            var result = action();
            return Task.FromResult(JsonSerializer.Serialize(result, IndentedJson));
        }
        catch (Exception ex)
        {
            return Task.FromResult($"Error: {ex.Message}");
        }
    }

    /// <summary>File-check + try-catch + JSON serialization for tools that return structured failure results.</summary>
    private static Task<string> ExecuteToolStructured<T>(string filePath, Func<T> action, Func<string, T> onError)
    {
        if (!File.Exists(filePath))
            return Task.FromResult(JsonSerializer.Serialize(onError($"File not found: {filePath}"), IndentedJson));
        try
        {
            var result = action();
            return Task.FromResult(JsonSerializer.Serialize(result, IndentedJson));
        }
        catch (Exception ex)
        {
            return Task.FromResult(JsonSerializer.Serialize(onError($"Error: {ex.Message}"), IndentedJson));
        }
    }

    /// <summary>List all slides in a PowerPoint presentation.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    [McpServerTool(Title = "List Slides", ReadOnly = true, Idempotent = true)]
    public Task<string> pptx_list_slides(string filePath) =>
        ExecuteToolJson(filePath, () => _service.GetSlides(filePath));

    /// <summary>List all available slide layouts in a PowerPoint presentation.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    [McpServerTool(Title = "List Layouts", ReadOnly = true, Idempotent = true)]
    public Task<string> pptx_list_layouts(string filePath) =>
        ExecuteToolJson(filePath, () => _service.GetLayouts(filePath));

    /// <summary>
    /// Update a named slide shape with replacement text while preserving the shape's existing formatting.
    /// Prefer shapeName from pptx_get_slide_content; placeholderIndex is a zero-based fallback across text-capable shapes on the slide.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideNumber">1-based slide number to update.</param>
    /// <param name="shapeName">Optional shape name to match exactly, ignoring case. When supplied and found, it takes precedence over placeholderIndex.</param>
    /// <param name="placeholderIndex">Optional zero-based fallback index across text-capable shapes on the slide.</param>
    /// <param name="newText">Replacement text for the target shape. Newlines create separate paragraphs.</param>
    [McpServerTool(Title = "Update Slide Data")]
    public partial Task<string> pptx_update_slide_data(string filePath, int slideNumber, string? shapeName = null, int? placeholderIndex = null, string newText = "") =>
        ExecuteToolStructured(filePath,
            () => _service.UpdateSlideData(filePath, slideNumber, shapeName, placeholderIndex, newText),
            error => new SlideDataUpdateResult(
                Success: false,
                SlideNumber: slideNumber,
                RequestedShapeName: shapeName,
                RequestedPlaceholderIndex: placeholderIndex,
                MatchedBy: null,
                ResolvedShapeName: null,
                ResolvedShapeIndex: null,
                ResolvedShapeId: null,
                PlaceholderType: null,
                LayoutPlaceholderIndex: null,
                PreviousText: null,
                NewText: newText,
                Message: error));

    /// <summary>
    /// Apply multiple named text updates across a presentation in a single open/save cycle.
    /// Each mutation targets a 1-based slide number and exact shape name, preserves formatting, and reports its own success or failure.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="mutations">Array of text mutations to apply. Each mutation must include slideNumber, shapeName, and newValue.</param>
    [McpServerTool(Title = "Batch Update")]
    public partial Task<string> pptx_batch_update(string filePath, BatchUpdateMutation[] mutations)
    {
        var requestedMutations = mutations ?? [];
        if (requestedMutations.Length == 0)
            return Task.FromResult(JsonSerializer.Serialize(new BatchUpdateResult(0, 0, 0, []), IndentedJson));

        return ExecuteToolStructured(filePath,
            () => _service.BatchUpdate(filePath, requestedMutations),
            error => new BatchUpdateResult(
                TotalMutations: requestedMutations.Length,
                SuccessCount: 0,
                FailureCount: requestedMutations.Length,
                Results: requestedMutations
                    .Select(m => new BatchUpdateMutationResult(
                        SlideNumber: m.SlideNumber,
                        ShapeName: m.ShapeName,
                        Success: false,
                        Error: error,
                        MatchedBy: null))
                    .ToArray()));
    }

    /// <summary>
    /// Set or replace the speaker notes on a slide. Pass append: true to add text after the existing notes.
    /// Use newlines (\n) in the notes string to create separate paragraphs.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideIndex">Zero-based index of the slide.</param>
    /// <param name="notes">Text to write as speaker notes. Use \n to separate paragraphs.</param>
    /// <param name="append">When true, appends to any existing notes instead of replacing them. Defaults to false.</param>
    [McpServerTool(Title = "Write Notes")]
    public Task<string> pptx_write_notes(string filePath, int slideIndex, string notes, bool append = false) =>
        ExecuteTool(filePath, () =>
        {
            _service.WriteNotes(filePath, slideIndex, notes, append);
            var mode = append ? "appended to" : "written to";
            return $"Notes {mode} slide {slideIndex} successfully.";
        });

    /// <summary>Insert an image onto a slide.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideIndex">Zero-based index of the slide.</param>
    /// <param name="imagePath">Absolute or relative path to the image file (.png, .jpg, .gif, .bmp).</param>
    /// <param name="x">Horizontal offset from the left edge of the slide in EMUs (English Metric Units). Default is 0.</param>
    /// <param name="y">Vertical offset from the top edge of the slide in EMUs. Default is 0.</param>
    /// <param name="width">Width of the image in EMUs. Default is 2743200 (~3 inches).</param>
    /// <param name="height">Height of the image in EMUs. Default is 2057400 (~2.25 inches).</param>
    [McpServerTool(Title = "Insert Image")]
    public Task<string> pptx_insert_image(
        string filePath,
        int slideIndex,
        string imagePath,
        long x = 0,
        long y = 0,
        long width = 2743200,
        long height = 2057400)
    {
        if (!File.Exists(filePath))
            return Task.FromResult($"Error: File not found: {filePath}");
        if (!File.Exists(imagePath))
            return Task.FromResult($"Error: Image file not found: {imagePath}");
        return ExecuteTool(filePath, () =>
        {
            _service.InsertImage(filePath, slideIndex, imagePath, x, y, width, height);
            return $"Image inserted successfully on slide {slideIndex}.";
        });
    }

    /// <summary>Get the raw XML of a specific slide.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideIndex">Zero-based index of the slide.</param>
    [McpServerTool(Title = "Get Slide XML", ReadOnly = true, Idempotent = true)]
    public Task<string> pptx_get_slide_xml(string filePath, int slideIndex) =>
        ExecuteTool(filePath, () => _service.GetSlideXml(filePath, slideIndex));

    /// <summary>
    /// Get structured content from a slide: all shapes with their type, position, size, and text.
    /// Returns a JSON object with slide dimensions and a shapes array. Each shape includes:
    /// ShapeType (Text, Picture, Table, Group, Connector, GraphicFrame), Name, position/size in EMUs,
    /// placeholder metadata when applicable, paragraph-level text for text shapes, and row/cell text for tables.
    /// Prefer this over pptx_get_slide_xml when you need to read or reason about slide content.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideIndex">Zero-based index of the slide.</param>
    [McpServerTool(Title = "Get Slide Content", ReadOnly = true, Idempotent = true)]
    public Task<string> pptx_get_slide_content(string filePath, int slideIndex) =>
        ExecuteToolJson(filePath, () => _service.GetSlideContent(filePath, slideIndex));

    /// <summary>
    /// Extract the highest-signal talking points from each slide in a PowerPoint presentation.
    /// The tool prioritizes body text and bullet-like content, filters common noise such as presenter notes labels
    /// and formatting-only text, and returns up to the requested number of points per slide.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="topN">Maximum number of talking points to return per slide. Defaults to 5.</param>
    [McpServerTool(Title = "Extract Talking Points", ReadOnly = true, Idempotent = true)]
    public Task<string> pptx_extract_talking_points(string filePath, int topN = 5) =>
        ExecuteToolJson(filePath, () => _service.ExtractTalkingPoints(filePath, topN));

    /// <summary>
    /// Export a PowerPoint presentation to markdown and save it as a .md file.
    /// The returned string is the generated markdown content with slide boundaries, headings,
    /// bullets, tables, and relative image references preserved for downstream processing.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="outputPath">Optional output path for the markdown file. Defaults to the presentation path with a .md extension.</param>
    [McpServerTool(Title = "Export Markdown", Idempotent = true)]
    public Task<string> pptx_export_markdown(string filePath, string? outputPath = null) =>
        ExecuteTool(filePath, () => _service.ExportMarkdown(filePath, outputPath).Markdown);

    /// <summary>Delete a slide from the presentation by its 1-based slide number.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideNumber">1-based number of the slide to delete.</param>
    [McpServerTool(Title = "Delete Slide")]
    public Task<string> pptx_delete_slide(string filePath, int slideNumber) =>
        ExecuteTool(filePath, () =>
        {
            _service.DeleteSlide(filePath, slideNumber);
            return $"Slide {slideNumber} deleted successfully.";
        });

    /// <summary>
    /// Insert a new table onto a slide. Pass column headers and data rows as arrays.
    /// Creates a DrawingML table (GraphicFrame) with proper PowerPoint-compatible structure.
    /// Position and size are specified in EMUs (English Metric Units). 1 inch = 914400 EMUs.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideNumber">1-based slide number to insert the table on.</param>
    /// <param name="headers">Array of column header strings. Determines the number of columns.</param>
    /// <param name="rows">Array of arrays, each containing cell values for one data row.</param>
    /// <param name="tableName">Optional name for the table. Defaults to "Table {id}".</param>
    /// <param name="x">Horizontal offset from the left edge in EMUs. Default is 914400 (1 inch).</param>
    /// <param name="y">Vertical offset from the top edge in EMUs. Default is 1371600 (1.5 inches).</param>
    /// <param name="width">Width of the table in EMUs. Default is 7315200 (~8 inches).</param>
    /// <param name="height">Height of the table in EMUs. Default is 1371600 (1.5 inches).</param>
    [McpServerTool(Title = "Insert Table")]
    public partial Task<string> pptx_insert_table(
        string filePath,
        int slideNumber,
        string[] headers,
        string[][] rows,
        string? tableName = null,
        long x = 914400,
        long y = 1371600,
        long width = 7315200,
        long height = 1371600) =>
        ExecuteToolStructured(filePath,
            () => _service.InsertTable(filePath, slideNumber, headers ?? [], rows ?? [], tableName, x, y, width, height),
            error => new TableInsertResult(
                Success: false,
                SlideNumber: slideNumber,
                TableName: tableName,
                TableShapeId: null,
                TableIndex: null,
                RowCount: 0,
                ColumnCount: 0,
                Message: error));

    /// <summary>
    /// Update cell values in an existing table on a slide.
    /// Locate the table by name (case-insensitive) or by zero-based index among tables on the slide.
    /// Each update targets a specific cell by zero-based row and column indices.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideNumber">1-based slide number containing the table.</param>
    /// <param name="updates">Array of cell updates. Each must include row (0-based), column (0-based), and value.</param>
    /// <param name="tableName">Optional table name to match (case-insensitive). Takes precedence over tableIndex.</param>
    /// <param name="tableIndex">Optional zero-based index among tables on the slide. Used when tableName is not provided.</param>
    [McpServerTool(Title = "Update Table")]
    public partial Task<string> pptx_update_table(
        string filePath,
        int slideNumber,
        TableCellUpdate[] updates,
        string? tableName = null,
        int? tableIndex = null)
    {
        if ((updates?.Length ?? 0) == 0)
        {
            var emptyResult = new TableUpdateResult(
                Success: false,
                SlideNumber: slideNumber,
                TableName: tableName,
                MatchedBy: null,
                CellsUpdated: 0,
                CellsSkipped: 0,
                Message: "No updates provided.");
            return Task.FromResult(JsonSerializer.Serialize(emptyResult, IndentedJson));
        }

        return ExecuteToolStructured(filePath,
            () => _service.UpdateTable(filePath, slideNumber, updates ?? [], tableName, tableIndex),
            error => new TableUpdateResult(
                Success: false,
                SlideNumber: slideNumber,
                TableName: tableName,
                MatchedBy: null,
                CellsUpdated: 0,
                CellsSkipped: 0,
                Message: error));
    }

    /// <summary>
    /// Replace an image in an existing picture shape on a slide.
    /// Target the picture by shape name (case-insensitive) or zero-based index among picture shapes on the slide.
    /// The replacement image inherits the existing shape geometry (position and size) so no manual EMU coordinates are needed.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideNumber">1-based slide number containing the picture to replace.</param>
    /// <param name="shapeName">Optional picture shape name to match (case-insensitive). Takes precedence over shapeIndex when both are provided.</param>
    /// <param name="shapeIndex">Optional zero-based index among picture shapes on the slide. Used as fallback when shapeName is not found or not provided.</param>
    /// <param name="imagePath">Absolute or relative path to the replacement image file (.png, .jpg, .jpeg, .svg).</param>
    /// <param name="altText">Optional alt text to set on the picture shape for accessibility.</param>
    [McpServerTool(Title = "Replace Image")]
    public partial Task<string> pptx_replace_image(
        string filePath,
        int slideNumber,
        string? shapeName = null,
        int? shapeIndex = null,
        string imagePath = "",
        string? altText = null)
    {
        ImageReplaceResult makeError(string message) => new(
            Success: false,
            SlideNumber: slideNumber,
            ShapeName: null,
            MatchedBy: null,
            PreviousImageContentType: null,
            NewImageContentType: null,
            AltText: altText,
            Message: message);

        if (!File.Exists(filePath))
            return Task.FromResult(JsonSerializer.Serialize(makeError($"File not found: {filePath}"), IndentedJson));
        if (!File.Exists(imagePath))
            return Task.FromResult(JsonSerializer.Serialize(makeError($"Image file not found: {imagePath}"), IndentedJson));

        try
        {
            var result = _service.ReplaceImage(filePath, slideNumber, shapeName, shapeIndex, imagePath, altText);
            return Task.FromResult(JsonSerializer.Serialize(result, IndentedJson));
        }
        catch (Exception ex)
        {
            return Task.FromResult(JsonSerializer.Serialize(makeError($"Error: {ex.Message}"), IndentedJson));
        }
    }
}


