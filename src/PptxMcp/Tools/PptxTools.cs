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
    public async Task<string> pptx_list_slides(string filePath)
    {
        if (!File.Exists(filePath))
            return $"Error: File not found: {filePath}";
        try
        {
            var slides = await _service.GetSlidesAsync(filePath);
            return JsonSerializer.Serialize(slides, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    /// <summary>List all available slide layouts in a PowerPoint presentation.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    [McpServerTool(Title = "List Layouts", ReadOnly = true, Idempotent = true)]
    public async Task<string> pptx_list_layouts(string filePath)
    {
        if (!File.Exists(filePath))
            return $"Error: File not found: {filePath}";
        try
        {
            var layouts = await _service.GetLayoutsAsync(filePath);
            return JsonSerializer.Serialize(layouts, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    /// <summary>Add a new slide to a PowerPoint presentation.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="layoutName">Optional name of the slide layout to use. Defaults to the first available layout.</param>
    [McpServerTool(Title = "Add Slide")]
    public async Task<string> pptx_add_slide(string filePath, string? layoutName = null)
    {
        if (!File.Exists(filePath))
            return $"Error: File not found: {filePath}";
        try
        {
            var newIndex = await _service.AddSlideAsync(filePath, layoutName);
            return $"Slide added successfully at index {newIndex}.";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    /// <summary>Update the text of a placeholder on a slide.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideIndex">Zero-based index of the slide to update.</param>
    /// <param name="placeholderIndex">Zero-based index of the placeholder on the slide.</param>
    /// <param name="text">New text content for the placeholder.</param>
    [McpServerTool(Title = "Update Text")]
    public async Task<string> pptx_update_text(string filePath, int slideIndex, int placeholderIndex, string text)
    {
        if (!File.Exists(filePath))
            return $"Error: File not found: {filePath}";
        try
        {
            await _service.UpdateTextPlaceholderAsync(filePath, slideIndex, placeholderIndex, text);
            return $"Placeholder {placeholderIndex} on slide {slideIndex} updated successfully.";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
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
    public async Task<string> pptx_insert_image(
        string filePath,
        int slideIndex,
        string imagePath,
        long x = 0,
        long y = 0,
        long width = 2743200,
        long height = 2057400)
    {
        if (!File.Exists(filePath))
            return $"Error: File not found: {filePath}";
        if (!File.Exists(imagePath))
            return $"Error: Image file not found: {imagePath}";
        try
        {
            await _service.InsertImageAsync(filePath, slideIndex, imagePath, x, y, width, height);
            return $"Image inserted successfully on slide {slideIndex}.";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    /// <summary>Get the raw XML of a specific slide.</summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideIndex">Zero-based index of the slide.</param>
    [McpServerTool(Title = "Get Slide XML", ReadOnly = true, Idempotent = true)]
    public async Task<string> pptx_get_slide_xml(string filePath, int slideIndex)
    {
        if (!File.Exists(filePath))
            return $"Error: File not found: {filePath}";
        try
        {
            return await _service.GetSlideXmlAsync(filePath, slideIndex);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }
}
