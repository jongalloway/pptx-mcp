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
    /// - Delete: Delete one or more slides by slide number.
    /// - DeleteAll: Delete all slides in the presentation.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="action">The slide management operation to perform: Add, AddFromLayout, Duplicate, Delete, or DeleteAll.</param>
    /// <param name="layoutName">Layout name. Required for AddFromLayout. Optional for Add (defaults to first available layout). Use pptx_list_layouts to discover available values.</param>
    /// <param name="slideNumber">1-based slide number to duplicate or delete. Required for Duplicate action. Optional for Delete when slideNumbers is not provided.</param>
    /// <param name="slideNumbers">Optional 1-based slide numbers to delete in a single operation. Used by Delete action.</param>
    /// <param name="placeholderValues">Optional placeholder text values keyed by semantic type with optional :index suffix (e.g. Title, Body:1, Picture:2). Used by AddFromLayout and Duplicate actions.</param>
    /// <param name="insertAt">Optional 1-based insertion position. Applies to AddFromLayout and Duplicate only. Defaults to end of deck for AddFromLayout, or after the source slide for Duplicate.</param>
    [McpServerTool(Title = "Manage Slides")]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["Add","AddFromLayout","Duplicate","Delete","DeleteAll"]""")]
    public partial Task<string> pptx_manage_slides(
        string filePath,
        ManageSlidesAction action,
        string? layoutName = null,
        int? slideNumber = null,
        int[]? slideNumbers = null,
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

            ManageSlidesAction.Delete => ExecuteToolStructured(filePath,
                () =>
                {
                    var requestedSlides = slideNumbers ?? [];
                    if (slideNumber is int singleSlideNumber)
                        requestedSlides = [.. requestedSlides, singleSlideNumber];

                    if (requestedSlides.Length == 0)
                        throw new ArgumentException("slideNumber or slideNumbers is required for the Delete action.");

                    var deletedCount = _service.DeleteSlides(filePath, requestedSlides);
                    return new DeleteSlidesResult(
                        Success: true,
                        DeletedCount: deletedCount,
                        DeletedSlideNumbers: requestedSlides.Distinct().Order().ToArray(),
                        Message: $"Deleted {deletedCount} slide(s).");
                },
                error => new DeleteSlidesResult(
                    Success: false,
                    DeletedCount: 0,
                    DeletedSlideNumbers: [],
                    Message: error)),

            ManageSlidesAction.DeleteAll => ExecuteToolStructured(filePath,
                () =>
                {
                    var deletedCount = _service.DeleteAllSlides(filePath);
                    return new DeleteSlidesResult(
                        Success: true,
                        DeletedCount: deletedCount,
                        DeletedSlideNumbers: [],
                        Message: $"Deleted all {deletedCount} slide(s).");
                },
                error => new DeleteSlidesResult(
                    Success: false,
                    DeletedCount: 0,
                    DeletedSlideNumbers: [],
                    Message: error)),

            _ => Task.FromResult(JsonSerializer.Serialize(
                new { Success = false, Message = $"Unknown action: {action}. Valid actions: Add, AddFromLayout, Duplicate, Delete, DeleteAll." },
                IndentedJson))
        };
    }
}
