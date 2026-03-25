using ModelContextProtocol.Protocol;
using PptxTools.Prompts;

namespace PptxTools.Tests.Prompts;

[Trait("Category", "Integration")]
public class PptxPromptsTests
{
    private readonly PptxPrompts _prompts = new();

    // --- RefreshQbrDeck ---

    [Fact]
    public void RefreshQbrDeck_ReturnsAtLeastOneMessage()
    {
        var messages = _prompts.RefreshQbrDeck("/path/to/qbr.pptx").ToList();
        Assert.NotEmpty(messages);
    }

    [Fact]
    public void RefreshQbrDeck_FirstMessageIsUserRole()
    {
        var messages = _prompts.RefreshQbrDeck("/path/to/qbr.pptx").ToList();
        Assert.Equal(Role.User, messages[0].Role);
    }

    [Fact]
    public void RefreshQbrDeck_ContainsFilePath()
    {
        const string path = "/my/deck.pptx";
        var messages = _prompts.RefreshQbrDeck(path).ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains(path, text);
    }

    [Fact]
    public void RefreshQbrDeck_ContainsMetricsSourceWhenProvided()
    {
        var messages = _prompts.RefreshQbrDeck("/deck.pptx", "Q3 sales data").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("Q3 sales data", text);
    }

    [Fact]
    public void RefreshQbrDeck_ContainsFallbackMetricsWhenNotProvided()
    {
        var messages = _prompts.RefreshQbrDeck("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("available metrics data", text);
    }

    [Fact]
    public void RefreshQbrDeck_MentionsUpdateSlideTool()
    {
        var messages = _prompts.RefreshQbrDeck("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_update_slide_data", text);
    }

    [Fact]
    public void RefreshQbrDeck_MentionsListSlidesTool()
    {
        var messages = _prompts.RefreshQbrDeck("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_list_slides", text);
    }

    // --- CreateAgendaSlide ---

    [Fact]
    public void CreateAgendaSlide_ReturnsAtLeastOneMessage()
    {
        var messages = _prompts.CreateAgendaSlide("/path/to/deck.pptx").ToList();
        Assert.NotEmpty(messages);
    }

    [Fact]
    public void CreateAgendaSlide_FirstMessageIsUserRole()
    {
        var messages = _prompts.CreateAgendaSlide("/path/to/deck.pptx").ToList();
        Assert.Equal(Role.User, messages[0].Role);
    }

    [Fact]
    public void CreateAgendaSlide_ContainsFilePath()
    {
        const string path = "/my/presentation.pptx";
        var messages = _prompts.CreateAgendaSlide(path).ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains(path, text);
    }

    [Fact]
    public void CreateAgendaSlide_DescribesAppendBehavior()
    {
        var messages = _prompts.CreateAgendaSlide("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("end of the deck", text);
    }

    [Fact]
    public void CreateAgendaSlide_MentionsManageSlidesTool()
    {
        var messages = _prompts.CreateAgendaSlide("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_manage_slides", text);
    }

    [Fact]
    public void CreateAgendaSlide_MentionsListLayoutsTool()
    {
        var messages = _prompts.CreateAgendaSlide("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_list_layouts", text);
    }

    // --- ReplaceKpiPlaceholders ---

    [Fact]
    public void ReplaceKpiPlaceholders_ReturnsAtLeastOneMessage()
    {
        var messages = _prompts.ReplaceKpiPlaceholders("/path/to/deck.pptx").ToList();
        Assert.NotEmpty(messages);
    }

    [Fact]
    public void ReplaceKpiPlaceholders_FirstMessageIsUserRole()
    {
        var messages = _prompts.ReplaceKpiPlaceholders("/path/to/deck.pptx").ToList();
        Assert.Equal(Role.User, messages[0].Role);
    }

    [Fact]
    public void ReplaceKpiPlaceholders_ContainsFilePath()
    {
        const string path = "/my/kpi-deck.pptx";
        var messages = _prompts.ReplaceKpiPlaceholders(path).ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains(path, text);
    }

    [Fact]
    public void ReplaceKpiPlaceholders_ContainsCustomPatternWhenProvided()
    {
        var messages = _prompts.ReplaceKpiPlaceholders("/deck.pptx", "{{METRIC}}").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("{{METRIC}}", text);
    }

    [Fact]
    public void ReplaceKpiPlaceholders_ContainsDefaultPatternWhenNotProvided()
    {
        var messages = _prompts.ReplaceKpiPlaceholders("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("TBD", text);
    }

    [Fact]
    public void ReplaceKpiPlaceholders_MentionsUpdateSlideTool()
    {
        var messages = _prompts.ReplaceKpiPlaceholders("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_update_slide_data", text);
    }

    [Fact]
    public void ReplaceKpiPlaceholders_MentionsGetSlideContentTool()
    {
        var messages = _prompts.ReplaceKpiPlaceholders("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_get_slide_content", text);
    }

    // --- BatchUpdateFromCsv ---

    [Fact]
    public void BatchUpdateFromCsv_ReturnsAtLeastOneMessage()
    {
        var messages = _prompts.BatchUpdateFromCsv("/path/to/deck.pptx", "/path/to/data.csv").ToList();
        Assert.NotEmpty(messages);
    }

    [Fact]
    public void BatchUpdateFromCsv_FirstMessageIsUserRole()
    {
        var messages = _prompts.BatchUpdateFromCsv("/path/to/deck.pptx", "/path/to/data.csv").ToList();
        Assert.Equal(Role.User, messages[0].Role);
    }

    [Fact]
    public void BatchUpdateFromCsv_ContainsFilePath()
    {
        const string path = "/my/presentation.pptx";
        var messages = _prompts.BatchUpdateFromCsv(path, "/data.csv").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains(path, text);
    }

    [Fact]
    public void BatchUpdateFromCsv_ContainsCsvPath()
    {
        const string csvPath = "/my/data/metrics.csv";
        var messages = _prompts.BatchUpdateFromCsv("/deck.pptx", csvPath).ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains(csvPath, text);
    }

    [Fact]
    public void BatchUpdateFromCsv_MentionsBatchUpdateTool()
    {
        var messages = _prompts.BatchUpdateFromCsv("/deck.pptx", "/data.csv").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_batch_update", text);
    }

    [Fact]
    public void BatchUpdateFromCsv_MentionsListSlidesTool()
    {
        var messages = _prompts.BatchUpdateFromCsv("/deck.pptx", "/data.csv").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_list_slides", text);
    }

    [Fact]
    public void BatchUpdateFromCsv_MentionsGetSlideContentTool()
    {
        var messages = _prompts.BatchUpdateFromCsv("/deck.pptx", "/data.csv").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_get_slide_content", text);
    }

    // --- ExtractForBlog ---

    [Fact]
    public void ExtractForBlog_ReturnsAtLeastOneMessage()
    {
        var messages = _prompts.ExtractForBlog("/path/to/deck.pptx").ToList();
        Assert.NotEmpty(messages);
    }

    [Fact]
    public void ExtractForBlog_FirstMessageIsUserRole()
    {
        var messages = _prompts.ExtractForBlog("/path/to/deck.pptx").ToList();
        Assert.Equal(Role.User, messages[0].Role);
    }

    [Fact]
    public void ExtractForBlog_ContainsFilePath()
    {
        const string path = "/my/slides.pptx";
        var messages = _prompts.ExtractForBlog(path).ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains(path, text);
    }

    [Fact]
    public void ExtractForBlog_WithNullFormat_UsesMarkdownDefault()
    {
        var messages = _prompts.ExtractForBlog("/deck.pptx", null).ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("markdown", text);
    }

    [Fact]
    public void ExtractForBlog_WithMarkdownFormat_IncludesFormatInText()
    {
        var messages = _prompts.ExtractForBlog("/deck.pptx", "markdown").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("markdown", text);
    }

    [Fact]
    public void ExtractForBlog_WithHtmlFormat_IncludesFormatInText()
    {
        var messages = _prompts.ExtractForBlog("/deck.pptx", "html").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("html", text);
    }

    [Fact]
    public void ExtractForBlog_MentionsExportMarkdownTool()
    {
        var messages = _prompts.ExtractForBlog("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_export_markdown", text);
    }

    [Fact]
    public void ExtractForBlog_MentionsExtractTalkingPointsTool()
    {
        var messages = _prompts.ExtractForBlog("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_extract_talking_points", text);
    }

    [Fact]
    public void ExtractForBlog_MentionsListSlidesTool()
    {
        var messages = _prompts.ExtractForBlog("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_list_slides", text);
    }

    // --- CreateSpeakerNotesOutline ---

    [Fact]
    public void CreateSpeakerNotesOutline_ReturnsAtLeastOneMessage()
    {
        var messages = _prompts.CreateSpeakerNotesOutline("/path/to/deck.pptx").ToList();
        Assert.NotEmpty(messages);
    }

    [Fact]
    public void CreateSpeakerNotesOutline_FirstMessageIsUserRole()
    {
        var messages = _prompts.CreateSpeakerNotesOutline("/path/to/deck.pptx").ToList();
        Assert.Equal(Role.User, messages[0].Role);
    }

    [Fact]
    public void CreateSpeakerNotesOutline_ContainsFilePath()
    {
        const string path = "/my/presentation.pptx";
        var messages = _prompts.CreateSpeakerNotesOutline(path).ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains(path, text);
    }

    [Fact]
    public void CreateSpeakerNotesOutline_WithNullStyle_UsesBulletPointsDefault()
    {
        var messages = _prompts.CreateSpeakerNotesOutline("/deck.pptx", null).ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("bullet-points", text);
    }

    [Fact]
    public void CreateSpeakerNotesOutline_WithBulletPointsStyle_IncludesStyleInText()
    {
        var messages = _prompts.CreateSpeakerNotesOutline("/deck.pptx", "bullet-points").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("bullet-points", text);
        Assert.Contains("concise bullet points", text);
    }

    [Fact]
    public void CreateSpeakerNotesOutline_WithNarrativeStyle_IncludesStyleInText()
    {
        var messages = _prompts.CreateSpeakerNotesOutline("/deck.pptx", "narrative").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("narrative", text);
        Assert.Contains("flowing narrative", text);
    }

    [Fact]
    public void CreateSpeakerNotesOutline_WithTimingCuesStyle_IncludesStyleInText()
    {
        var messages = _prompts.CreateSpeakerNotesOutline("/deck.pptx", "timing-cues").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("timing-cues", text);
        Assert.Contains("timing estimates", text);
    }

    [Fact]
    public void CreateSpeakerNotesOutline_MentionsWriteNotesTool()
    {
        var messages = _prompts.CreateSpeakerNotesOutline("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_write_notes", text);
    }

    [Fact]
    public void CreateSpeakerNotesOutline_MentionsExtractTalkingPointsTool()
    {
        var messages = _prompts.CreateSpeakerNotesOutline("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_extract_talking_points", text);
    }

    [Fact]
    public void CreateSpeakerNotesOutline_MentionsListSlidesTool()
    {
        var messages = _prompts.CreateSpeakerNotesOutline("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_list_slides", text);
    }

    // --- OptimizeForWeb ---

    [Fact]
    public void OptimizeForWeb_ReturnsAtLeastOneMessage()
    {
        var messages = _prompts.OptimizeForWeb("/path/to/deck.pptx").ToList();
        Assert.NotEmpty(messages);
    }

    [Fact]
    public void OptimizeForWeb_FirstMessageIsUserRole()
    {
        var messages = _prompts.OptimizeForWeb("/path/to/deck.pptx").ToList();
        Assert.Equal(Role.User, messages[0].Role);
    }

    [Fact]
    public void OptimizeForWeb_ContainsFilePath()
    {
        const string path = "/my/large-deck.pptx";
        var messages = _prompts.OptimizeForWeb(path).ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains(path, text);
    }

    [Fact]
    public void OptimizeForWeb_WithNullTargetSize_UsesGeneralGuidance()
    {
        var messages = _prompts.OptimizeForWeb("/deck.pptx", null).ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("Optimize as much as possible", text);
    }

    [Fact]
    public void OptimizeForWeb_WithTargetSize_IncludesTargetInText()
    {
        var messages = _prompts.OptimizeForWeb("/deck.pptx", 5.0).ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("5", text);
        Assert.Contains("MB", text);
    }

    [Fact]
    public void OptimizeForWeb_MentionsAnalyzeFileSizeTool()
    {
        var messages = _prompts.OptimizeForWeb("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_analyze_file_size", text);
    }

    [Fact]
    public void OptimizeForWeb_MentionsOptimizeImagesTool()
    {
        var messages = _prompts.OptimizeForWeb("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_optimize_images", text);
    }

    [Fact]
    public void OptimizeForWeb_MentionsManageMediaTool()
    {
        var messages = _prompts.OptimizeForWeb("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_manage_media", text);
    }

    [Fact]
    public void OptimizeForWeb_MentionsManageLayoutsTool()
    {
        var messages = _prompts.OptimizeForWeb("/deck.pptx").ToList();
        var text = GetMessageText(messages[0]);
        Assert.Contains("pptx_manage_layouts", text);
    }

    // --- Helpers ---

    private static string GetMessageText(PromptMessage message)
    {
        var textBlock = Assert.IsType<TextContentBlock>(message.Content);
        return textBlock.Text;
    }
}
