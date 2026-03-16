namespace PptxMcp.Tests.Tools;

public class PptxExportMarkdownToolTests : IDisposable
{
    private readonly PresentationService _service = new();
    private readonly PptxTools _tools;
    private readonly List<string> _tempArtifacts = [];

    public PptxExportMarkdownToolTests()
    {
        _tools = new PptxTools(_service);
    }

    public void Dispose()
    {
        foreach (var artifact in _tempArtifacts.OrderByDescending(path => path.Length))
        {
            if (File.Exists(artifact))
                File.Delete(artifact);
            else if (Directory.Exists(artifact))
                Directory.Delete(artifact, recursive: true);
        }
    }

    [Fact]
    public async Task pptx_export_markdown_ReturnsMarkdownAndWritesFile()
    {
        var path = CreatePresentation(
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

    private string CreatePresentation(params TestSlideDefinition[] slides)
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        _tempArtifacts.Add(path);
        TestPptxHelper.CreatePresentation(path, slides);
        return path;
    }

    private string CreateOutputPath()
    {
        var directory = Path.Join(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(directory);
        _tempArtifacts.Add(directory);
        return Path.Join(directory, "tool-output.md");
    }
}
