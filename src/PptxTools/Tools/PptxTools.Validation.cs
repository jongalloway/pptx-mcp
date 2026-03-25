using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxTools.Models;

namespace PptxTools.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Validate a PowerPoint presentation for structural issues and integrity problems.
    /// Available actions:
    /// - Validate: Check for duplicate shape IDs, missing image references, orphaned relationships,
    ///   broken hyperlink targets, and missing required XML elements. Returns a structured report
    ///   with severity levels (Error, Warning, Info) and repair recommendations.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="action">The validation operation to perform: Validate.</param>
    /// <param name="slideNumber">Optional 1-based slide number to restrict validation to a single slide.</param>
    [McpServerTool(Title = "Validate Presentation", ReadOnly = true, Idempotent = true)]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["Validate"]""")]
    public partial Task<string> pxtx_validate_presentation(
        string filePath,
        ValidationAction action,
        int? slideNumber = null)
    {
        return action switch
        {
            ValidationAction.Validate => ExecuteToolStructured(filePath,
                () => _service.ValidatePresentation(filePath, slideNumber),
                error => new ValidationResult(
                    Success: false,
                    Action: "Validate",
                    IssueCount: 0,
                    ErrorCount: 0,
                    WarningCount: 0,
                    InfoCount: 0,
                    Issues: [],
                    Message: error)),

            _ => Task.FromResult(JsonSerializer.Serialize(
                new { Success = false, Message = $"Unknown action: {action}. Valid actions: Validate." },
                IndentedJson))
        };
    }
}
