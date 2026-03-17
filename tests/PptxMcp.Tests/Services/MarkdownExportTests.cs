using DocumentFormat.OpenXml.Presentation;
using System.Text.RegularExpressions;

namespace PptxMcp.Tests.Services;

[Trait("Category", "Unit")]
public class MarkdownExportTests : PptxTestBase
{

    [Fact]
    public void ExportMarkdown_WritesMarkdownFileAndSlideHeading()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Quarterly Review",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Paragraphs = ["Executive summary"],
                        PlaceholderType = PlaceholderValues.Body
                    }
                ]
            });
        var outputPath = CreateOutputPath();

        var export = Service.ExportMarkdown(path, outputPath);

        Assert.Equal(Path.GetFullPath(outputPath), export.OutputPath);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("# Quarterly Review", export.Markdown);
        Assert.Contains("## Slide 1: Quarterly Review", export.Markdown);
        Assert.Contains("Executive summary", File.ReadAllText(outputPath));
    }

    [Fact]
    public void ExportMarkdown_PreservesNestedBulletIndentation()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Release Plan",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Ship the MCP tool", IsBullet = true },
                            new TestParagraphDefinition { Text = "Export markdown", IsBullet = true, Level = 1 },
                            new TestParagraphDefinition { Text = "Verify output", IsBullet = true, Level = 2 }
                        ]
                    }
                ]
            });

        var export = Service.ExportMarkdown(path, CreateOutputPath());

        Assert.Contains("- Ship the MCP tool", export.Markdown);
        Assert.Contains("  - Export markdown", export.Markdown);
        Assert.Contains("    - Verify output", export.Markdown);
    }

    [Fact]
    public void ExportMarkdown_FormatsSubtitleAsHeading()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Kickoff",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Paragraphs = ["Agenda"],
                        PlaceholderType = PlaceholderValues.SubTitle
                    },
                    new TestTextShapeDefinition
                    {
                        Paragraphs = ["Discuss scope and owners"]
                    }
                ]
            });

        var export = Service.ExportMarkdown(path, CreateOutputPath());

        Assert.Contains("### Agenda", export.Markdown);
        Assert.Contains("Discuss scope and owners", export.Markdown);
    }

    [Fact]
    public void ExportMarkdown_ExportsImageWithRelativeReference()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Architecture",
                IncludeImage = true
            });
        var outputPath = CreateOutputPath();

        var export = Service.ExportMarkdown(path, outputPath);
        var imageMatch = Regex.Match(export.Markdown, @"!\[[^\]]+\]\(([^)]+)\)");
        var imageDirectory = Path.Join(
            Path.GetDirectoryName(outputPath)!,
            $"{Path.GetFileNameWithoutExtension(outputPath)}_images");

        Assert.True(imageMatch.Success);
        Assert.Equal(1, export.ImageCount);
        Assert.True(Directory.Exists(imageDirectory));
        Assert.Single(Directory.GetFiles(imageDirectory));
        Assert.DoesNotContain("\\", imageMatch.Groups[1].Value);
    }

    [Fact]
    public void ExportMarkdown_FormatsTablesAsMarkdown()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Metrics",
                Tables =
                [
                    new TestTableDefinition
                    {
                        Rows =
                        [
                            ["Metric", "Value"],
                            ["Users", "1200"],
                            ["Revenue", "$1M"]
                        ]
                    }
                ]
            });

        var export = Service.ExportMarkdown(path, CreateOutputPath());

        Assert.Contains("| Metric | Value |", export.Markdown);
        Assert.Contains("| --- | --- |", export.Markdown);
        Assert.Contains("| Users | 1200 |", export.Markdown);
        Assert.Contains("| Revenue | $1M |", export.Markdown);
    }

    [Fact]
    public void ExportMarkdown_DefaultOutputPath_WritesMarkdownAlongsideSource()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Default Path Test" });
        var expectedOutputPath = Path.ChangeExtension(path, ".md");
        TrackTempFile(expectedOutputPath);

        var export = Service.ExportMarkdown(path);

        Assert.Equal(Path.GetFullPath(expectedOutputPath), export.OutputPath);
        Assert.True(File.Exists(expectedOutputPath));
        Assert.Contains("# Default Path Test", export.Markdown);
    }

    [Fact]
    public void ExportMarkdown_RealWorldStyleDeck_RendersSlidesInOrder()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Launch Plan",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Paragraphs = ["Goals"],
                        PlaceholderType = PlaceholderValues.SubTitle
                    },
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Finalize markdown export", IsBullet = true },
                            new TestParagraphDefinition { Text = "Validate on sample decks", IsBullet = true }
                        ]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Architecture",
                IncludeImage = true
            },
            new TestSlideDefinition
            {
                TitleText = "Metrics",
                Tables =
                [
                    new TestTableDefinition
                    {
                        Rows =
                        [
                            ["Metric", "Status"],
                            ["Build", "Passing"],
                            ["Tests", "Passing"]
                        ]
                    }
                ]
            });

        var export = Service.ExportMarkdown(path, CreateOutputPath());

        var slide1 = export.Markdown.IndexOf("## Slide 1: Launch Plan", StringComparison.Ordinal);
        var slide2 = export.Markdown.IndexOf("## Slide 2: Architecture", StringComparison.Ordinal);
        var slide3 = export.Markdown.IndexOf("## Slide 3: Metrics", StringComparison.Ordinal);

        Assert.True(slide1 >= 0 && slide2 > slide1 && slide3 > slide2);
        Assert.Contains("### Goals", export.Markdown);
        Assert.Contains("- Finalize markdown export", export.Markdown);
        Assert.Contains("![Picture", export.Markdown);
        Assert.Contains("| Metric | Status |", export.Markdown);
    }

    private string CreateOutputPath()
    {
        var directory = Path.Join(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(directory);
        TrackTempFile(directory);
        return Path.Join(directory, "presentation.md");
    }
}
