using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace PptxTools.Resources;

public sealed partial class PptxResources
{
    /// <summary>
    /// Browse all images in a PowerPoint presentation as a JSON resource.
    /// Returns an array of image objects with slide number, shape name, format, content type,
    /// relationship ID, and dimensions.
    /// </summary>
    [McpServerResource(UriTemplate = "pptx://{file}/images", Name = "images", Title = "Images", MimeType = "application/json")]
    public TextResourceContents GetImages(string file)
    {
        var decodedPath = Uri.UnescapeDataString(file);
        var canonicalFile = Uri.EscapeDataString(decodedPath);
        string json;
        if (!File.Exists(decodedPath))
        {
            json = JsonSerializer.Serialize(new { error = $"File not found: {decodedPath}" },
                new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            var images = _service.GetImageInfos(decodedPath);
            json = JsonSerializer.Serialize(images, new JsonSerializerOptions { WriteIndented = true });
        }
        return new TextResourceContents
        {
            Uri = $"pptx://{canonicalFile}/images",
            MimeType = "application/json",
            Text = json
        };
    }

    /// <summary>
    /// Browse presentation-level metadata as a JSON resource.
    /// Returns title, author, dates, subject, keywords, description, last modified by, and slide count.
    /// </summary>
    [McpServerResource(UriTemplate = "pptx://{file}/presentation-metadata", Name = "presentation-metadata", Title = "Presentation Metadata", MimeType = "application/json")]
    public TextResourceContents GetPresentationMetadata(string file)
    {
        var decodedPath = Uri.UnescapeDataString(file);
        var canonicalFile = Uri.EscapeDataString(decodedPath);
        string json;
        if (!File.Exists(decodedPath))
        {
            json = JsonSerializer.Serialize(new { error = $"File not found: {decodedPath}" },
                new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            var metadata = _service.GetPresentationMetadata(decodedPath);
            json = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });
        }
        return new TextResourceContents
        {
            Uri = $"pptx://{canonicalFile}/presentation-metadata",
            MimeType = "application/json",
            Text = json
        };
    }

    /// <summary>
    /// Browse all tables in a PowerPoint presentation as a JSON resource.
    /// Returns an array of table objects with slide number, table name, row/column counts,
    /// and header row text.
    /// </summary>
    [McpServerResource(UriTemplate = "pptx://{file}/tables", Name = "tables", Title = "Tables", MimeType = "application/json")]
    public TextResourceContents GetTables(string file)
    {
        var decodedPath = Uri.UnescapeDataString(file);
        var canonicalFile = Uri.EscapeDataString(decodedPath);
        string json;
        if (!File.Exists(decodedPath))
        {
            json = JsonSerializer.Serialize(new { error = $"File not found: {decodedPath}" },
                new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            var allSlides = _service.GetAllSlideContents(decodedPath);
            var tables = new List<object>();
            foreach (var slide in allSlides)
            {
                foreach (var shape in slide.Shapes.Where(s => s.ShapeType == "Table" && s.TableRows is not null))
                {
                    tables.Add(new
                    {
                        SlideNumber = slide.SlideIndex + 1,
                        TableName = shape.Name,
                        RowCount = shape.TableRows!.Count,
                        ColumnCount = shape.TableRows!.Count > 0 ? shape.TableRows![0].Count : 0,
                        HeaderRow = shape.TableRows!.Count > 0 ? shape.TableRows![0] : (IReadOnlyList<string>)Array.Empty<string>()
                    });
                }
            }
            json = JsonSerializer.Serialize(tables, new JsonSerializerOptions { WriteIndented = true });
        }
        return new TextResourceContents
        {
            Uri = $"pptx://{canonicalFile}/tables",
            MimeType = "application/json",
            Text = json
        };
    }

    /// <summary>
    /// Browse all speaker notes in a PowerPoint presentation as a JSON resource.
    /// Returns an array of note objects with slide number, title, and notes text
    /// for slides that have speaker notes.
    /// </summary>
    [McpServerResource(UriTemplate = "pptx://{file}/notes", Name = "notes", Title = "Speaker Notes", MimeType = "application/json")]
    public TextResourceContents GetNotes(string file)
    {
        var decodedPath = Uri.UnescapeDataString(file);
        var canonicalFile = Uri.EscapeDataString(decodedPath);
        string json;
        if (!File.Exists(decodedPath))
        {
            json = JsonSerializer.Serialize(new { error = $"File not found: {decodedPath}" },
                new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            var slides = _service.GetSlides(decodedPath);
            var notes = slides
                .Where(s => s.Notes is not null)
                .Select(s => new
                {
                    SlideNumber = s.Index + 1,
                    s.Title,
                    s.Notes
                })
                .ToList();
            json = JsonSerializer.Serialize(notes, new JsonSerializerOptions { WriteIndented = true });
        }
        return new TextResourceContents
        {
            Uri = $"pptx://{canonicalFile}/notes",
            MimeType = "application/json",
            Text = json
        };
    }
}
