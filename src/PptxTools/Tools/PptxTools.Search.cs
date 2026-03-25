using ModelContextProtocol.Server;
using PptxTools.Models;

namespace PptxTools.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Search for text across all slides in a PowerPoint presentation.
    /// Finds shapes containing the specified text and returns match details with surrounding context.
    /// Searches both text shapes and table cell content.
    /// Set useRegex to true to treat searchText as a regular expression pattern.
    /// Optionally filter to a single slide with slideNumber.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="searchText">Text string or regex pattern to search for.</param>
    /// <param name="caseSensitive">When false (default), matches ignore case. Ignored when useRegex is true.</param>
    /// <param name="slideNumber">Optional 1-based slide number to limit the search to a single slide.</param>
    /// <param name="useRegex">When true, treat searchText as a regular expression pattern. Defaults to false.</param>
    [McpServerTool(Title = "Search Text", ReadOnly = true, Idempotent = true)]
    public partial Task<string> pptx_search_text(
        string filePath,
        string searchText,
        bool caseSensitive = false,
        int? slideNumber = null,
        bool useRegex = false) =>
        ExecuteToolStructured(filePath,
            () => useRegex
                ? _service.SearchByRegex(filePath, searchText, slideNumber)
                : _service.SearchText(filePath, searchText, caseSensitive, slideNumber),
            error => new TextSearchResult(
                Success: false,
                Matches: [],
                TotalMatches: 0,
                SlidesSearched: 0,
                Message: error));

    /// <summary>
    /// Find shapes with no text content in a PowerPoint presentation.
    /// Identifies empty text shapes and tables — useful for detecting unfilled template placeholders
    /// or cleaning up unused shapes. Only considers shapes that could hold text (Text and Table types).
    /// Optionally filter to a single slide with slideNumber.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="slideNumber">Optional 1-based slide number to limit the search to a single slide.</param>
    [McpServerTool(Title = "Find Empty Shapes", ReadOnly = true, Idempotent = true)]
    public partial Task<string> pptx_find_empty_shapes(
        string filePath,
        int? slideNumber = null) =>
        ExecuteToolStructured(filePath,
            () => _service.FindEmptyShapes(filePath, slideNumber),
            error => new EmptyShapeResult(
                Success: false,
                EmptyShapes: [],
                TotalFound: 0,
                SlidesSearched: 0,
                Message: error));
}
