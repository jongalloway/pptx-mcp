namespace PptxMcp.Tests.Tools;

public class PptxExportMarkdownToolTests : PptxTestBase
{
    private readonly PptxTools _tools;

    public PptxExportMarkdownToolTests()
    {
        _tools = new PptxTools(Service);
    }

    [Fact]
    public async Task pptx_export_markdown_ReturnsMarkdownAndWritesFile()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Tool Output",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Paragraphs = ["Generate markdown for agents"],
                        PlaceholderType = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body
                    }
                ]
            });
        var outputPath = CreateOutputPath();

        var result = await _tools.pptx_export_markdown(path, outputPath);

        Assert.Contains("# Tool Output", result);
        Assert.Contains("## Slide 1: Tool Output", result);
        Assert.True(File.Exists(outputPath));
        Assert.Equal(result, File.ReadAllText(outputPath));
    }

    [Fact]
    public async Task pptx_export_markdown_FileNotFound_ReturnsError()
    {
        var result = await _tools.pptx_export_markdown("C:\\does-not-exist\\missing.pptx");

        Assert.StartsWith("Error:", result);
    }

    private string CreateOutputPath()
    {
        var directory = Path.Join(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(directory);
        TrackTempFile(directory);
        return Path.Join(directory, "tool-output.md");
    }
}
