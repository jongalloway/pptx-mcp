using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxTools.Models;

namespace PptxTools.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Create slides in a PowerPoint presentation.
    /// Available actions:
    /// - Add: Add a blank slide with an optional layout name.
    /// - AddFromLayout: Create a slide from a named layout and optionally populate placeholders.
    /// - Duplicate: Clone an existing slide with optional placeholder overrides.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="action">The slide creation operation to perform: Add, AddFromLayout, or Duplicate.</param>
    /// <param name="layoutName">Layout name. Required for AddFromLayout. Optional for Add (defaults to first available layout). Use pptx_list_layouts to discover available values.</param>
    /// <param name="slideNumber">1-based slide number to duplicate. Required for Duplicate action.</param>
    /// <param name="placeholderValues">Optional placeholder text values keyed by semantic type with optional :index suffix (e.g. Title, Body:1, Picture:2). Used by AddFromLayout and Duplicate actions.</param>
    /// <param name="insertAt">Optional 1-based insertion position. Applies to AddFromLayout and Duplicate only. Defaults to end of deck for AddFromLayout, or after the source slide for Duplicate.</param>
    [McpServerTool(Title = "Manage Slides")]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["Add","AddFromLayout","Duplicate"]""")]
    public partial Task<string> pptx_manage_slides(
        string filePath,
        ManageSlidesAction action,
        string? layoutName = null,
        int? slideNumber = null,
        Dictionary<string, string>? placeholderValues = null,
        int? insertAt = null)
    {
        return action switch
        {
            ManageSlidesAction.Add => ExecuteToolStructured(filePath,
                () =>
                {
                    var newIndex = _service.AddSlide(filePath, layoutName);
                    return new AddSlideResult(
                        Success: true,
                        SlideNumber: newIndex + 1,
                        LayoutName: layoutName,
                        Message: $"Added slide {newIndex + 1}.");
                },
                error => new AddSlideResult(
                    Success: false,
                    SlideNumber: null,
                    LayoutName: layoutName,
                    Message: error)),

            ManageSlidesAction.AddFromLayout => ExecuteToolStructured(filePath,
                () =>
                {
                    if (string.IsNullOrWhiteSpace(layoutName))
                        throw new ArgumentException("layoutName is required for the AddFromLayout action.");
                    return _service.AddSlideFromLayout(filePath, layoutName, placeholderValues, insertAt);
                },
                error => new AddSlideFromLayoutResult(
                    Success: false,
                    SlideNumber: null,
                    LayoutName: layoutName,
                    PlaceholdersPopulated: 0,
                    Message: error)),

            ManageSlidesAction.Duplicate => ExecuteToolStructured(filePath,
                () =>
                {
                    if (slideNumber is null)
                        throw new ArgumentException("slideNumber is required for the Duplicate action.");
                    return _service.DuplicateSlide(filePath, slideNumber.Value, placeholderValues, insertAt);
                },
                error => new DuplicateSlideResult(
                    Success: false,
                    NewSlideNumber: null,
                    ShapesCopied: 0,
                    OverridesApplied: 0,
                    Message: error)),

            _ => Task.FromResult(JsonSerializer.Serialize(
                new { Success = false, Message = $"Unknown action: {action}. Valid actions: Add, AddFromLayout, Duplicate." },
                IndentedJson))
        };
    }
}
