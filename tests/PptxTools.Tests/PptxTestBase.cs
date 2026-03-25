namespace PptxTools.Tests;

public abstract class PptxTestBase : IDisposable
{
    protected readonly PresentationService Service = new();
    private readonly List<string> _tempArtifacts = [];

    protected string CreateMinimalPptx(string? titleText = "Test Slide")
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        _tempArtifacts.Add(path);
        TestPptxHelper.CreateMinimalPresentation(path, titleText);
        return path;
    }

    protected string CreatePptxWithSlides(params TestSlideDefinition[] slides)
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        _tempArtifacts.Add(path);
        TestPptxHelper.CreatePresentation(path, slides);
        return path;
    }

    protected void TrackTempFile(string path) => _tempArtifacts.Add(path);

    public virtual void Dispose()
    {
        foreach (var artifact in _tempArtifacts.OrderByDescending(p => p.Length))
        {
            if (File.Exists(artifact)) File.Delete(artifact);
            else if (Directory.Exists(artifact)) Directory.Delete(artifact, recursive: true);
        }
    }
}
