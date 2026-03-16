using System.Text.Json;
using ModelContextProtocol.Server;
using PptxMcp.Services;

namespace PptxMcp.Tools;

[McpServerToolType]
public sealed class PptxTools
{
    private readonly PresentationService _service;

    public PptxTools(PresentationService service)
    {
        _service = service;
    }

    /// <summary>List all slides in a PowerPoint presentation.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    [McpServerTool(Title = "List Slides", ReadOnly = true, Idempotent = true)]
    public Task<string> pptx_list_slides(string filePath)
    {
        if (!File.Exists(filePath))
            return Task.FromResult($"Error: File not found: {filePath}");
        try
        {
            var slides = _service.GetSlides(filePath);
            return Task.FromResult(JsonSerializer.Serialize(slides, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            return Task.FromResult($"Error: {ex.Message}");
        }
    }

    /// <summary>List all available slide layouts in a PowerPoint presentation.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    [McpServerTool(Title = "List Layouts", ReadOnly = true, Idempotent = true)]
    public Task<string> pptx_list_layouts(string filePath)
    {
        if (!File.Exists(filePath))
            return Task.FromResult($"Error: File not found: {filePath}");
        try
        {
            var layouts = _service.GetLayouts(filePath);
            return Task.FromResult(JsonSerializer.Serialize(layouts, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            return Task.FromResult($"Error: {ex.Message}");
        }
    }

    /// <summary>Add a new slide to a PowerPoint presentation.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="layoutName">Optional name of the slide layout to use. Defaults to the first available layout.</param>
    [McpServerTool(Title = "Add Slide")]
    public Task<string> pptx_add_slide(string filePath, string? layoutName = null)
    {
        if (!File.Exists(filePath))
            return Task.FromResult($"Error: File not found: {filePath}");
        try
        {
            var newIndex = _service.AddSlide(filePath, layoutName);
            return Task.FromResult($"Slide added successfully at index {newIndex}.");
        }
        catch (Exception ex)
        {
            return Task.FromResult($"Error: {ex.Message}");
        }
    }

    /// <summary>Update the text of a placeholder on a slide.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideIndex">Zero-based index of the slide to update.</param>
    /// <param name="placeholderIndex">Zero-based index of the placeholder on the slide.</param>
    /// <param name="text">New text content for the placeholder.</param>
    [McpServerTool(Title = "Update Text")]
    public Task<string> pptx_update_text(string filePath, int slideIndex, int placeholderIndex, string text)
    {
        if (!File.Exists(filePath))
            return Task.FromResult($"Error: File not found: {filePath}");
        try
        {
            _service.UpdateTextPlaceholder(filePath, slideIndex, placeholderIndex, text);
            return Task.FromResult($"Placeholder {placeholderIndex} on slide {slideIndex} updated successfully.");
        }
        catch (Exception ex)
        {
            return Task.FromResult($"Error: {ex.Message}");
        }
    }

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
        try
        {
            _service.InsertImage(filePath, slideIndex, imagePath, x, y, width, height);
            return Task.FromResult($"Image inserted successfully on slide {slideIndex}.");
        }
        catch (Exception ex)
        {
            return Task.FromResult($"Error: {ex.Message}");
        }
    }

    /// <summary>Get the raw XML of a specific slide.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideIndex">Zero-based index of the slide.</param>
    [McpServerTool(Title = "Get Slide XML", ReadOnly = true, Idempotent = true)]
    public Task<string> pptx_get_slide_xml(string filePath, int slideIndex)
    {
        if (!File.Exists(filePath))
            return Task.FromResult($"Error: File not found: {filePath}");
        try
        {
            return Task.FromResult(_service.GetSlideXml(filePath, slideIndex));
        }
        catch (Exception ex)
        {
            return Task.FromResult($"Error: {ex.Message}");
        }
    }

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
    public Task<string> pptx_get_slide_content(string filePath, int slideIndex)
    {
        if (!File.Exists(filePath))
            return Task.FromResult($"Error: File not found: {filePath}");
        try
        {
            var content = _service.GetSlideContent(filePath, slideIndex);
            return Task.FromResult(JsonSerializer.Serialize(content, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            return Task.FromResult($"Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Export a PowerPoint presentation to markdown and save it as a .md file.
    /// The returned string is the generated markdown content with slide boundaries, headings,
    /// bullets, tables, and relative image references preserved for downstream processing.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="outputPath">Optional output path for the markdown file. Defaults to the presentation path with a .md extension.</param>
    [McpServerTool(Title = "Export Markdown", Idempotent = true)]
    public Task<string> pptx_export_markdown(string filePath, string? outputPath = null)
    {
        if (!File.Exists(filePath))
            return Task.FromResult($"Error: File not found: {filePath}");
        try
        {
            var export = _service.ExportMarkdown(filePath, outputPath);
            return Task.FromResult(export.Markdown);
        }
        catch (Exception ex)
        {
            return Task.FromResult($"Error: {ex.Message}");
        }
    }
}
