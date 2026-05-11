using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxTools.Models;

namespace PptxTools.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Inspect the theme colors and fonts from a PowerPoint presentation.
    /// Returns the color scheme (Dark1/2, Light1/2, Accent1–6, Hyperlink, FollowedHyperlink)
    /// and font scheme (MajorFont for headings, MinorFont for body text) for each slide master.
    /// Useful for template analysis, branding verification, and CSS theme generation.
    /// Colors are returned as #RRGGBB hex strings.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    [McpServerTool(Title = "Inspect Theme", ReadOnly = true, Idempotent = true)]
    public partial Task<string> pptx_inspect_theme(string filePath)
    {
        return ExecuteToolStructured(filePath,
            () => _service.GetThemeInfo(filePath),
            error => new ThemeInfoResult(
                Success: false,
                ThemeCount: 0,
                Themes: [],
                Message: error));
    }
}
