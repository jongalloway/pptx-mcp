using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxMcp.Models;

namespace PptxMcp.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Change slide order in a PowerPoint presentation.
    /// Available actions:
    /// - Move: Move a single slide to a different position.
    /// - Reorder: Specify the complete new slide sequence as a 1-based array.
    /// For example, to reverse a 3-slide deck, use Reorder with newOrder [3, 2, 1].
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="action">The ordering operation to perform: Move or Reorder.</param>
    /// <param name="slideNumber">1-based slide number to move. Required for Move action.</param>
    /// <param name="targetPosition">1-based target position. Required for Move action.</param>
    /// <param name="newOrder">Complete new slide order as a 1-based array. Required for Reorder action. Must be a permutation of 1..n where n is the slide count.</param>
    [McpServerTool(Title = "Reorder Slides")]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["Move","Reorder"]""")]
    public partial Task<string> pptx_reorder_slides(
        string filePath,
        ReorderSlidesAction action,
        int? slideNumber = null,
        int? targetPosition = null,
        int[]? newOrder = null)
    {
        return action switch
        {
            ReorderSlidesAction.Move => ExecuteToolStructured(filePath,
                () =>
                {
                    if (slideNumber is null)
                        throw new ArgumentException("slideNumber is required for the Move action.");
                    if (targetPosition is null)
                        throw new ArgumentException("targetPosition is required for the Move action.");
                    _service.MoveSlide(filePath, slideNumber.Value, targetPosition.Value);
                    return new SlideOrderResult(
                        Success: true,
                        Action: "Move",
                        Message: $"Slide {slideNumber} moved to position {targetPosition}.");
                },
                error => new SlideOrderResult(
                    Success: false,
                    Action: "Move",
                    Message: error)),

            ReorderSlidesAction.Reorder => ExecuteToolStructured(filePath,
                () =>
                {
                    if (newOrder is null || newOrder.Length == 0)
                        throw new ArgumentException("newOrder is required for the Reorder action.");
                    _service.ReorderSlides(filePath, newOrder);
                    return new SlideOrderResult(
                        Success: true,
                        Action: "Reorder",
                        Message: "Slides reordered successfully.");
                },
                error => new SlideOrderResult(
                    Success: false,
                    Action: "Reorder",
                    Message: error)),

            _ => Task.FromResult(JsonSerializer.Serialize(
                new { Success = false, Message = $"Unknown action: {action}. Valid actions: Move, Reorder." },
                IndentedJson))
        };
    }
}
