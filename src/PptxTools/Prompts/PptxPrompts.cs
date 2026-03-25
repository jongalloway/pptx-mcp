using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace PptxTools.Prompts;

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
    /// Note: <c>pptx_manage_slides</c> with the Add action appends the new slide at the end of the deck.
    /// </summary>
    /// <param name="filePath">Absolute path to the .pptx file.</param>
    [McpServerPrompt(Name = "create-agenda-slide", Title = "Create Agenda Slide")]
    public IEnumerable<PromptMessage> CreateAgendaSlide(string filePath)
    {
        yield return new PromptMessage
        {
            Role = Role.User,
            Content = new TextContentBlock
            {
                Text = $"""
                    I need to add an agenda slide to the presentation at "{filePath}".

                    Please follow these steps:
                    1. Use pptx_list_slides to get all current slide titles.
                    2. Use pptx_list_layouts to find an appropriate layout (prefer "Title and Content" layouts for agenda slides).
                    3. Use pptx_manage_slides with action Add and the chosen layout to add the new slide at the end of the deck.
                    4. Use pptx_get_slide_content on the new slide to find the title and body placeholder names.
                    5. Use pptx_update_slide_data to set the title to "Agenda" and the body to a bulleted list of the other slide titles.
                    6. Confirm the agenda slide was created and summarize its content.

                    Note: pptx_manage_slides with the Add action always appends slides at the end of the deck. If a specific position is required, use pptx_reorder_slides with the Move action afterwards.
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

    /// <summary>
    /// Generate a step-by-step workflow for batch updating slide content from a CSV file
    /// by mapping CSV columns to shapes and applying all changes in one pass.
    /// </summary>
    /// <param name="filePath">Absolute path to the .pptx file to update.</param>
    /// <param name="csvPath">Absolute path to the CSV file containing the update data.</param>
    [McpServerPrompt(Name = "batch-update-from-csv", Title = "Batch Update from CSV")]
    public IEnumerable<PromptMessage> BatchUpdateFromCsv(string filePath, string csvPath)
    {
        yield return new PromptMessage
        {
            Role = Role.User,
            Content = new TextContentBlock
            {
                Text = $"""
                    I need to batch update the presentation at "{filePath}" using data from the CSV file at "{csvPath}".

                    Please follow these steps:
                    1. Read and parse the CSV file to understand its structure (columns and rows).
                    2. Use pptx_list_slides to see all slides in the deck.
                    3. Use pptx_get_slide_content on relevant slides to discover shape names that should be populated from CSV columns.
                    4. Map CSV columns to slide numbers and shape names (e.g., column "Revenue" → Slide 3, shape "Revenue Value").
                    5. Build a batch update request with all mutations (one per CSV row or per shape mapping as appropriate).
                    6. Use pptx_batch_update to apply all changes in a single operation.
                    7. Confirm all updates and summarize how many shapes were updated and which CSV data was applied.
                    """
            }
        };
    }

    /// <summary>
    /// Generate a step-by-step workflow for extracting presentation content and restructuring
    /// it as a blog post with headings, body text, and image descriptions.
    /// </summary>
    /// <param name="filePath">Absolute path to the .pptx file to extract from.</param>
    /// <param name="format">Output format for the blog post: "markdown" (default) or "html".</param>
    [McpServerPrompt(Name = "extract-for-blog", Title = "Extract Content for Blog Post")]
    public IEnumerable<PromptMessage> ExtractForBlog(string filePath, string? format = null)
    {
        var outputFormat = string.IsNullOrWhiteSpace(format) || format.Equals("markdown", StringComparison.OrdinalIgnoreCase)
            ? "markdown"
            : "html";

        yield return new PromptMessage
        {
            Role = Role.User,
            Content = new TextContentBlock
            {
                Text = $"""
                    I need to extract content from the presentation at "{filePath}" and restructure it as a blog post in {outputFormat} format.

                    Please follow these steps:
                    1. Use pptx_list_slides to get all slide titles and understand the flow.
                    2. Use pptx_get_slide_content on each slide to extract the full content (titles, body text, bullets, tables).
                    3. Use pptx_extract_talking_points to identify key points from each slide that should be emphasized in the blog post.
                    4. Use pptx_export_markdown to get the baseline markdown export with embedded image references.
                    5. Restructure the exported content into a blog post format:
                       - Use slide titles as section headings (## or <h2>)
                       - Convert bullet points to flowing paragraphs or keep as lists where appropriate
                       - Include inline descriptions of any images (e.g., "Figure 1: Sales trend graph showing...")
                       - Combine related slides into cohesive sections
                    6. Output the final blog post in {outputFormat} format with proper structure and readability.
                    """
            }
        };
    }

    /// <summary>
    /// Generate a step-by-step workflow for analyzing slide content and creating
    /// structured speaker notes for each slide.
    /// </summary>
    /// <param name="filePath">Absolute path to the .pptx file.</param>
    /// <param name="style">Speaker notes style: "bullet-points" (default), "narrative", or "timing-cues".</param>
    [McpServerPrompt(Name = "create-speaker-notes-outline", Title = "Create Speaker Notes Outline")]
    public IEnumerable<PromptMessage> CreateSpeakerNotesOutline(string filePath, string? style = null)
    {
        var notesStyle = string.IsNullOrWhiteSpace(style) ? "bullet-points" : style.ToLowerInvariant();
        var styleDescription = notesStyle switch
        {
            "narrative" => "flowing narrative paragraphs that tell the story of each slide",
            "timing-cues" => "bullet points with timing estimates and transition cues (e.g., '(30 sec)', 'PAUSE HERE')",
            _ => "concise bullet points summarizing key talking points"
        };

        yield return new PromptMessage
        {
            Role = Role.User,
            Content = new TextContentBlock
            {
                Text = $"""
                    I need to create speaker notes for the presentation at "{filePath}" in {notesStyle} style.
                    The notes should be {styleDescription}.

                    Please follow these steps:
                    1. Use pptx_list_slides to see all slides in the deck.
                    2. Use pptx_get_slide_content on each slide to understand the slide content and structure.
                    3. Use pptx_extract_talking_points to identify the key points that should be covered when presenting each slide.
                    4. For each slide, generate speaker notes in {notesStyle} style that expand on the slide content and provide guidance for the presenter.
                    5. Use pptx_write_notes to add the generated speaker notes to each slide.
                    6. Confirm all notes were added and provide a summary showing the first line of notes for each slide.
                    """
            }
        };
    }

    /// <summary>
    /// Generate a step-by-step workflow for optimizing a presentation for web/email distribution
    /// by reducing file size through image optimization and cleanup.
    /// </summary>
    /// <param name="filePath">Absolute path to the .pptx file to optimize.</param>
    /// <param name="targetSizeMb">Target file size in megabytes (optional). If specified, optimize until this target is reached.</param>
    [McpServerPrompt(Name = "optimize-for-web", Title = "Optimize Presentation for Web/Email")]
    public IEnumerable<PromptMessage> OptimizeForWeb(string filePath, double? targetSizeMb = null)
    {
        var targetGuidance = targetSizeMb.HasValue
            ? $"The target file size is {targetSizeMb.Value} MB. Continue optimizing until this target is reached or no further optimizations are possible."
            : "Optimize as much as possible while maintaining visual quality.";

        yield return new PromptMessage
        {
            Role = Role.User,
            Content = new TextContentBlock
            {
                Text = $"""
                    I need to optimize the presentation at "{filePath}" for web or email distribution by reducing the file size.
                    {targetGuidance}

                    Please follow these steps:
                    1. Use pptx_analyze_file_size to understand the current file size breakdown and identify the largest contributors (media, relationships, text).
                    2. Use pptx_manage_media with the Analyze action to get detailed information about all embedded images, videos, and audio files.
                    3. Use pptx_optimize_images to compress images while maintaining acceptable visual quality for web/email viewing.
                    4. Use pptx_manage_layouts with the Find action to identify unused slide layouts, then use the Remove action to clean them up.
                    5. Use pptx_manage_media with the Deduplicate action to remove duplicate media files that may be embedded multiple times.
                    6. Use pptx_analyze_file_size again to measure the impact of optimizations.
                    7. Summarize the optimization results: original size, final size, reduction percentage, and which optimization steps had the most impact.
                    """
            }
        };
    }
}
