namespace PptxTools.Models;

public sealed record MarkdownExportResult(
    string OutputPath,
    string Markdown,
    int SlideCount,
    int ImageCount);
