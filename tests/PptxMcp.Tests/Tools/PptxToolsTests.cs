using DocumentFormat.OpenXml.Packaging;

namespace PptxMcp.Tests.Tools;

public class PptxToolsTests : IDisposable
{
    private readonly PresentationService _service = new();
    private readonly PptxTools _tools;
    private readonly List<string> _tempFiles = new();

    public PptxToolsTests()
    {
        _tools = new PptxTools(_service);
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            if (File.Exists(f)) File.Delete(f);
    }

    private string CreateTempPptx()
    {
        var path = Path.GetTempFileName() + ".pptx";
        _tempFiles.Add(path);
        TestPptxHelper.CreateMinimalPresentation(path, "Test Slide");
        return path;
    }

    [Fact]
    public async Task pptx_list_slides_ReturnsJson()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_list_slides(path);
        Assert.Contains("Index", result);
    }

    [Fact]
    public async Task pptx_list_slides_FileNotFound_ReturnsError()
    {
        var result = await _tools.pptx_list_slides("/nonexistent/path/file.pptx");
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_list_layouts_ReturnsJson()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_list_layouts(path);
        Assert.Contains("Name", result);
    }

    [Fact]
    public async Task pptx_list_layouts_FileNotFound_ReturnsError()
    {
        var result = await _tools.pptx_list_layouts("/nonexistent/path/file.pptx");
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task pptx_add_slide_ReturnsSuccessMessage()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_add_slide(path);
        Assert.Contains("successfully", result);
    }

    [Fact]
    public async Task pptx_update_text_ReturnsSuccessMessage()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_update_text(path, 0, 0, "New Text");
        Assert.Contains("successfully", result);
    }

    [Fact]
    public async Task pptx_get_slide_xml_ReturnsXml()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_get_slide_xml(path, 0);
        Assert.Contains("sld", result);
    }

    [Fact]
    public async Task pptx_insert_image_FileNotFound_ReturnsError()
    {
        var path = CreateTempPptx();
        var result = await _tools.pptx_insert_image(path, 0, "/nonexistent/image.png");
        Assert.StartsWith("Error:", result);
    }
}
