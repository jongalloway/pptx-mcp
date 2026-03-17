using System.Text.Json;
using ModelContextProtocol.Server;
using PptxMcp.Models;

namespace PptxMcp.Tools;

public sealed partial class PptxTools
{
    /// <summary>
    /// Create a new slide from a named layout and optionally populate placeholders using semantic keys like Title or Body:1.
    /// The new slide keeps its relationship to the selected layout so PowerPoint can continue to inherit layout and master geometry.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="layoutName">Exact layout name to use. Use pptx_list_layouts to discover available values.</param>
    /// <param name="placeholderValues">Optional placeholder text values keyed by semantic placeholder type, optionally with a :index suffix such as Title, Body:1, or Picture:2.</param>
    /// <param name="insertAt">Optional 1-based insertion position. Defaults to appending the slide at the end of the deck.</param>
    [McpServerTool(Title = "Add Slide From Layout")]
    public partial Task<string> pptx_add_slide_from_layout(string filePath, string layoutName, Dictionary<string, string>? placeholderValues = null, int? insertAt = null)
    {
        if (!File.Exists(filePath))
        {
            var missingFileResult = new AddSlideFromLayoutResult(
                Success: false,
                SlideNumber: null,
                LayoutName: layoutName,
                PlaceholdersPopulated: 0,
                Message: $"File not found: {filePath}");

            return Task.FromResult(JsonSerializer.Serialize(missingFileResult, new JsonSerializerOptions { WriteIndented = true }));
        }

        try
        {
            var result = _service.AddSlideFromLayout(filePath, layoutName, placeholderValues, insertAt);
            return Task.FromResult(JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            var failureResult = new AddSlideFromLayoutResult(
                Success: false,
                SlideNumber: null,
                LayoutName: layoutName,
                PlaceholdersPopulated: 0,
                Message: $"Error: {ex.Message}");

            return Task.FromResult(JsonSerializer.Serialize(failureResult, new JsonSerializerOptions { WriteIndented = true }));
        }
    }

    /// <summary>
    /// Duplicate a slide and optionally override placeholders using semantic keys like Title or Body:1.
    /// Related slide parts such as images are cloned so the duplicated slide is independent from the source slide.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideNumber">1-based slide number to duplicate.</param>
    /// <param name="placeholderOverrides">Optional placeholder text overrides keyed by semantic placeholder type, optionally with a :index suffix.</param>
    /// <param name="insertAt">Optional 1-based insertion position. Defaults to inserting immediately after the source slide.</param>
    [McpServerTool(Title = "Duplicate Slide")]
    public partial Task<string> pptx_duplicate_slide(string filePath, int slideNumber, Dictionary<string, string>? placeholderOverrides = null, int? insertAt = null)
    {
        if (!File.Exists(filePath))
        {
            var missingFileResult = new DuplicateSlideResult(
                Success: false,
                NewSlideNumber: null,
                ShapesCopied: 0,
                OverridesApplied: 0,
                Message: $"File not found: {filePath}");

            return Task.FromResult(JsonSerializer.Serialize(missingFileResult, new JsonSerializerOptions { WriteIndented = true }));
        }

        try
        {
            var result = _service.DuplicateSlide(filePath, slideNumber, placeholderOverrides, insertAt);
            return Task.FromResult(JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            var failureResult = new DuplicateSlideResult(
                Success: false,
                NewSlideNumber: null,
                ShapesCopied: 0,
                OverridesApplied: 0,
                Message: $"Error: {ex.Message}");

            return Task.FromResult(JsonSerializer.Serialize(failureResult, new JsonSerializerOptions { WriteIndented = true }));
        }
    }
}
