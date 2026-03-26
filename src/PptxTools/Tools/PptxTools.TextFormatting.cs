using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxTools.Models;

namespace PptxTools.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Read or apply text formatting in a PowerPoint presentation.
    /// Available actions:
    /// - Get: Read all text formatting properties (font, size, bold, italic, underline, color, alignment) from shapes on a slide.
    /// - Apply: Apply font styling to all text runs in a target shape.
    /// Use Get first to inspect current formatting, then Apply to change it.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="action">The text formatting operation to perform: Get or Apply.</param>
    /// <param name="slideNumber">1-based slide number. Optional for Get (returns all slides when omitted), required for Apply.</param>
    /// <param name="shapeName">Shape name to filter/target. Optional for Get, required for Apply. Case-insensitive match.</param>
    /// <param name="fontFamily">Font family to apply (e.g. "Arial", "Calibri"). Apply action only.</param>
    /// <param name="fontSize">Font size in points (e.g. 12, 24.5). Apply action only.</param>
    /// <param name="bold">Set bold on or off. Apply action only.</param>
    /// <param name="italic">Set italic on or off. Apply action only.</param>
    /// <param name="underline">Set underline on or off. Apply action only.</param>
    /// <param name="color">Hex RGB color string (e.g. "#FF0000" for red). Apply action only.</param>
    /// <param name="alignment">Paragraph alignment: Left, Center, Right, or Justify. Apply action only.</param>
    [McpServerTool(Title = "Text Formatting")]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["Get","Apply"]""")]
    public partial Task<string> pptx_manage_text_formatting(
        string filePath,
        TextFormattingAction action,
        int? slideNumber = null,
        string? shapeName = null,
        string? fontFamily = null,
        double? fontSize = null,
        bool? bold = null,
        bool? italic = null,
        bool? underline = null,
        string? color = null,
        string? alignment = null)
    {
        return action switch
        {
            TextFormattingAction.Get => ExecuteToolStructured(filePath,
                () => _service.GetTextFormatting(filePath, slideNumber, shapeName),
                error => new TextFormattingResult(
                    Success: false,
                    Action: "Get",
                    SlideNumber: slideNumber,
                    ShapeName: shapeName,
                    FormattingCount: 0,
                    Formattings: [],
                    Message: error)),

            TextFormattingAction.Apply => ExecuteToolStructured(filePath,
                () =>
                {
                    if (slideNumber is null)
                        throw new ArgumentException("slideNumber is required for the Apply action.");
                    if (string.IsNullOrWhiteSpace(shapeName))
                        throw new ArgumentException("shapeName is required for the Apply action.");
                    return _service.ApplyTextFormatting(filePath, slideNumber.Value, shapeName!,
                        fontFamily, fontSize, bold, italic, underline, color, alignment);
                },
                error => new TextFormattingResult(
                    Success: false,
                    Action: "Apply",
                    SlideNumber: slideNumber,
                    ShapeName: shapeName,
                    FormattingCount: 0,
                    Formattings: [],
                    Message: error)),

            _ => Task.FromResult(JsonSerializer.Serialize(
                new { Success = false, Message = $"Unknown action: {action}. Valid actions: Get, Apply." },
                IndentedJson))
        };
    }
}
