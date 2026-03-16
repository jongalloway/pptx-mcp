using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using PptxMcp.Services;

namespace PptxMcp.Resources;

/// <summary>
/// MCP resources for browsing PowerPoint presentation state.
/// Resource URIs use <c>pptx://{file}/...</c> where <c>{file}</c> is the URL-encoded
/// absolute path to the .pptx file.
/// </summary>
[McpServerResourceType]
public sealed class PptxResources
{
    private readonly PresentationService _service;

    public PptxResources(PresentationService service)
    {
        _service = service;
    }

    /// <summary>
    /// Browse all slides in a PowerPoint presentation as a JSON resource.
    /// Returns an array of slide objects with index, title, notes, and placeholder count.
    /// </summary>
    [McpServerResource(UriTemplate = "pptx://{file}/slides", Name = "slides", Title = "Slides", MimeType = "application/json")]
    public TextResourceContents GetSlides(string file)
    {
        var decodedPath = Uri.UnescapeDataString(file);
        string json;
        if (!File.Exists(decodedPath))
        {
            json = JsonSerializer.Serialize(new { error = $"File not found: {decodedPath}" },
                new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            var slides = _service.GetSlides(decodedPath);
            json = JsonSerializer.Serialize(slides, new JsonSerializerOptions { WriteIndented = true });
        }
        return new TextResourceContents
        {
            Uri = $"pptx://{file}/slides",
            MimeType = "application/json",
            Text = json
        };
    }

    /// <summary>
    /// Browse all available slide layouts in a PowerPoint presentation as a JSON resource.
    /// Returns an array of layout objects with index and name.
    /// </summary>
    [McpServerResource(UriTemplate = "pptx://{file}/layouts", Name = "layouts", Title = "Layouts", MimeType = "application/json")]
    public TextResourceContents GetLayouts(string file)
    {
        var decodedPath = Uri.UnescapeDataString(file);
        string json;
        if (!File.Exists(decodedPath))
        {
            json = JsonSerializer.Serialize(new { error = $"File not found: {decodedPath}" },
                new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            var layouts = _service.GetLayouts(decodedPath);
            json = JsonSerializer.Serialize(layouts, new JsonSerializerOptions { WriteIndented = true });
        }
        return new TextResourceContents
        {
            Uri = $"pptx://{file}/layouts",
            MimeType = "application/json",
            Text = json
        };
    }

    /// <summary>
    /// Browse a map of all named shapes across every slide in a PowerPoint presentation as a JSON resource.
    /// Returns an object keyed by slide index, each containing an array of shapes with name, type,
    /// placeholder type, and current text — useful for targeting shapes by name with pptx_update_slide_data.
    /// </summary>
    [McpServerResource(UriTemplate = "pptx://{file}/shape-map", Name = "shape-map", Title = "Shape Map", MimeType = "application/json")]
    public TextResourceContents GetShapeMap(string file)
    {
        var decodedPath = Uri.UnescapeDataString(file);
        string json;
        if (!File.Exists(decodedPath))
        {
            json = JsonSerializer.Serialize(new { error = $"File not found: {decodedPath}" },
                new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            var slides = _service.GetSlides(decodedPath);
            var shapeMap = new Dictionary<string, object>();
            for (int i = 0; i < slides.Count; i++)
            {
                var content = _service.GetSlideContent(decodedPath, i);
                shapeMap[$"slide{i}"] = content.Shapes.Select(s => new
                {
                    s.Name,
                    s.ShapeType,
                    s.PlaceholderType,
                    s.Text,
                    s.IsPlaceholder
                }).ToList();
            }
            json = JsonSerializer.Serialize(shapeMap, new JsonSerializerOptions { WriteIndented = true });
        }
        return new TextResourceContents
        {
            Uri = $"pptx://{file}/shape-map",
            MimeType = "application/json",
            Text = json
        };
    }
}
