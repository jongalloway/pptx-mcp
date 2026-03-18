using ModelContextProtocol.Protocol;
using PptxMcp.Prompts;

namespace PptxMcp.Tests.Prompts;

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

    // --- Helpers ---

    private static string GetMessageText(PromptMessage message)
    {
        var textBlock = Assert.IsType<TextContentBlock>(message.Content);
        return textBlock.Text;
    }
}
