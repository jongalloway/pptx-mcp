using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace PptxMcp.Prompts;

/// <summary>
/// MCP prompt templates for common PowerPoint workflow scenarios.
/// </summary>
[McpServerPromptType]
public sealed class PptxPrompts
{
    /// <summary>
    /// Generate a step-by-step workflow for refreshing a QBR (Quarterly Business Review) deck
    /// by pulling live metrics from a data source and updating the named shapes.
    /// </summary>
    /// <param name="filePath">Absolute path to the QBR .pptx file to refresh.</param>
    /// <param name="metricsSource">Description of the metrics data source (e.g., "last week's sales report", "Q3 dashboard CSV").</param>
    [McpServerPrompt(Name = "refresh-qbr-deck", Title = "Refresh QBR Deck from Metrics")]
    public IEnumerable<PromptMessage> RefreshQbrDeck(string filePath, string? metricsSource = null)
    {
        var source = string.IsNullOrWhiteSpace(metricsSource)
            ? "the available metrics data"
            : metricsSource;

        yield return new PromptMessage
        {
            Role = Role.User,
            Content = new TextContentBlock
            {
                Text = $"""
                    I need to refresh the QBR deck at "{filePath}" with the latest metrics from {source}.

                    Please follow these steps:
                    1. Use pptx_list_slides to see all slides in the deck.
                    2. Use pptx_get_slide_content on each slide to identify named shapes that hold KPI values (look for shapes with names like "Revenue Value", "Growth Rate", "Target", etc.).
                    3. Fetch or summarize the relevant metrics from {source}.
                    4. For each KPI shape, use pptx_update_slide_data with the shape name and the new value to update it while preserving formatting.
                    5. Confirm all updates and summarize which shapes were changed and what the new values are.
                    """
            }
        };
    }

    /// <summary>
    /// Generate a step-by-step workflow for adding a new agenda slide that lists the
    /// current slide titles as agenda items.
    /// </summary>
    /// <param name="filePath">Absolute path to the .pptx file.</param>
    /// <param name="insertAfterSlide">1-based slide number after which to insert the agenda slide. Defaults to inserting after slide 1 (the title slide).</param>
    [McpServerPrompt(Name = "create-agenda-slide", Title = "Create Agenda Slide")]
    public IEnumerable<PromptMessage> CreateAgendaSlide(string filePath, int insertAfterSlide = 1)
    {
        yield return new PromptMessage
        {
            Role = Role.User,
            Content = new TextContentBlock
            {
                Text = $"""
                    I need to add an agenda slide to the presentation at "{filePath}", inserted after slide {insertAfterSlide}.

                    Please follow these steps:
                    1. Use pptx_list_slides to get all current slide titles.
                    2. Use pptx_list_layouts to find an appropriate layout (prefer "Title and Content" or "Two Content" layouts for agenda slides).
                    3. Use pptx_add_slide with the chosen layout to add the new slide. Note the index of the new slide.
                    4. Use pptx_get_slide_content on the new slide to find the title and body placeholder names.
                    5. Use pptx_update_slide_data to set the title to "Agenda" and the body to a bulleted list of the other slide titles.
                    6. Confirm the agenda slide was created and summarize its content.

                    The agenda items should include all slides except the title slide and the new agenda slide itself.
                    """
            }
        };
    }

    /// <summary>
    /// Generate a step-by-step workflow for finding and replacing all KPI placeholder
    /// text tokens in a presentation with actual values.
    /// </summary>
    /// <param name="filePath">Absolute path to the .pptx file.</param>
    /// <param name="placeholderPattern">Pattern used to identify placeholder tokens in the slides, e.g. "{{KPI}}" or "TBD". Defaults to common patterns like "{{...}}", "[VALUE]", and "TBD".</param>
    [McpServerPrompt(Name = "replace-kpi-placeholders", Title = "Replace KPI Placeholders")]
    public IEnumerable<PromptMessage> ReplaceKpiPlaceholders(string filePath, string? placeholderPattern = null)
    {
        var pattern = string.IsNullOrWhiteSpace(placeholderPattern)
            ? "placeholder tokens like {{KPI_NAME}}, [VALUE], or TBD"
            : $"\"{placeholderPattern}\" tokens";

        yield return new PromptMessage
        {
            Role = Role.User,
            Content = new TextContentBlock
            {
                Text = $"""
                    I need to find and replace all KPI placeholder values in the presentation at "{filePath}".
                    I'm looking for {pattern}.

                    Please follow these steps:
                    1. Use pptx_get_slide_content on each slide to inspect all shape text.
                    2. Identify every shape that contains {pattern}.
                    3. For each shape with a placeholder, ask me for the real value (or infer it from context if a data source is available).
                    4. Use pptx_update_slide_data with the shape name to replace the placeholder text with the real value, preserving the shape's existing formatting.
                    5. After all replacements, summarize which shapes were updated, what was replaced, and what the new values are.

                    Start by scanning all slides and listing every shape that contains placeholder text, then we'll go through them one by one.
                    """
            }
        };
    }
}
