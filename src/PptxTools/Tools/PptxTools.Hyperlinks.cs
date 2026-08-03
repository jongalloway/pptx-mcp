using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxTools.Models;

namespace PptxTools.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Manage hyperlinks in a PowerPoint presentation.
    /// Available actions:
    /// - Get: List all hyperlinks in the presentation, optionally filtered to a specific slide (read-only).
    /// - Add: Add an external hyperlink to a shape on a slide.
    /// - Update: Change the URL (and optional tooltip) of an existing hyperlink on a shape.
    /// - Remove: Remove all hyperlinks from a shape.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="action">The hyperlink operation to perform: Get, Add, Update, or Remove.</param>
    /// <param name="slideNumber">1-based slide number. Required for Add, Update, and Remove. Optional filter for Get.</param>
    /// <param name="shapeName">Name of the target shape. Required for Add, Update, and Remove.</param>
    /// <param name="url">Hyperlink URL. Required for Add and Update. Supports http://, https://, and mailto: URLs.</param>
    /// <param name="tooltip">Optional tooltip text displayed on hover. Used by Add and Update.</param>
    [McpServerTool(Title = "Manage Hyperlinks")]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["Get","Add","Update","Remove"]""")]
    public partial Task<string> pptx_manage_hyperlinks(
        string filePath,
        HyperlinkAction action,
        int? slideNumber = null,
        string? shapeName = null,
        string? url = null,
        string? tooltip = null)
    {
        return action switch
        {
            HyperlinkAction.Get => ExecuteToolStructured(filePath,
                () =>
                {
                    var hyperlinks = _service.GetHyperlinks(filePath, slideNumber);
                    return new HyperlinkResult(
                        Success: true,
                        Action: "Get",
                        SlideNumber: slideNumber,
                        ShapeName: null,
                        Url: null,
                        HyperlinkCount: hyperlinks.Count,
                        Hyperlinks: hyperlinks,
                        Message: hyperlinks.Count == 0
                            ? "No hyperlinks found."
                            : $"Found {hyperlinks.Count} hyperlink(s).");
                },
                error => new HyperlinkResult(
                    Success: false, Action: "Get", SlideNumber: slideNumber,
                    ShapeName: null, Url: null, HyperlinkCount: 0, Hyperlinks: null,
                    Message: error)),

            HyperlinkAction.Add => ExecuteToolStructured(filePath,
                () =>
                {
                    if (slideNumber is null)
                        throw new ArgumentException("slideNumber is required for the Add action.");
                    if (string.IsNullOrWhiteSpace(shapeName))
                        throw new ArgumentException("shapeName is required for the Add action.");
                    if (string.IsNullOrWhiteSpace(url))
                        throw new ArgumentException("url is required for the Add action.");
                    return _service.AddHyperlink(filePath, slideNumber.Value, shapeName, url, tooltip);
                },
                error => new HyperlinkResult(
                    Success: false, Action: "Add", SlideNumber: slideNumber,
                    ShapeName: shapeName, Url: url, HyperlinkCount: 0, Hyperlinks: null,
                    Message: error)),

            HyperlinkAction.Update => ExecuteToolStructured(filePath,
                () =>
                {
                    if (slideNumber is null)
                        throw new ArgumentException("slideNumber is required for the Update action.");
                    if (string.IsNullOrWhiteSpace(shapeName))
                        throw new ArgumentException("shapeName is required for the Update action.");
                    if (string.IsNullOrWhiteSpace(url))
                        throw new ArgumentException("url is required for the Update action.");
                    return _service.UpdateHyperlink(filePath, slideNumber.Value, shapeName, url, tooltip);
                },
                error => new HyperlinkResult(
                    Success: false, Action: "Update", SlideNumber: slideNumber,
                    ShapeName: shapeName, Url: url, HyperlinkCount: 0, Hyperlinks: null,
                    Message: error)),

            HyperlinkAction.Remove => ExecuteToolStructured(filePath,
                () =>
                {
                    if (slideNumber is null)
                        throw new ArgumentException("slideNumber is required for the Remove action.");
                    if (string.IsNullOrWhiteSpace(shapeName))
                        throw new ArgumentException("shapeName is required for the Remove action.");
                    return _service.RemoveHyperlink(filePath, slideNumber.Value, shapeName);
                },
                error => new HyperlinkResult(
                    Success: false, Action: "Remove", SlideNumber: slideNumber,
                    ShapeName: shapeName, Url: null, HyperlinkCount: 0, Hyperlinks: null,
                    Message: error)),

            _ => Task.FromResult(JsonSerializer.Serialize(
                new { Success = false, Message = $"Unknown action: {action}. Valid actions: Get, Add, Update, Remove." },
                IndentedJson))
        };
    }
}
