using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Tests.Tools;

[Trait("Category", "E2E")]
public class PptxPhase1E2eTests : PptxTestBase
{
    private readonly PptxTools _tools;

    public PptxPhase1E2eTests()
    {
        _tools = new PptxTools(Service);
    }

    [Fact]
    public async Task Phase1Tools_ProductUpdateDeck_ExtractsTalkingPointsAndExportsMarkdown()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "FY26 Launch Review",
                SpeakerNotesText = "Mention the pilot customers only after slide 2.",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.SubTitle,
                        Paragraphs = ["Executive summary"]
                    },
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Launch completed in 3 regions", IsBullet = true },
                            new TestParagraphDefinition { Text = "Error rate dropped below 1%", IsBullet = true },
                            new TestParagraphDefinition { Text = "Support volume stayed flat", IsBullet = true }
                        ]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Adoption Highlights",
                SpeakerNotesText = "Do not read the backup metric in notes.",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Enterprise expansion in APAC", IsBullet = true },
                            new TestParagraphDefinition { Text = "2 lighthouse customers went live", IsBullet = true },
                            new TestParagraphDefinition { Text = "Renewal intent reached 94%", IsBullet = true }
                        ]
                    }
                ],
                Tables =
                [
                    new TestTableDefinition
                    {
                        Rows =
                        [
                            ["Segment", "Status"],
                            ["SMB", "Stable"],
                            ["Enterprise", "Growing"]
                        ]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Architecture Snapshot",
                SpeakerNotesText = "Diagram shows internal service names.",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Thin MCP tools delegate to PresentationService", IsBullet = true },
                            new TestParagraphDefinition { Text = "OpenXML parsing stays centralized", IsBullet = true }
                        ]
                    }
                ],
                IncludeImage = true
            },
            new TestSlideDefinition
            {
                TitleText = "Next Steps",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Publish markdown workflows to customers", IsBullet = true },
                            new TestParagraphDefinition { Text = "Track Unicode and image-only edge cases", IsBullet = true }
                        ]
                    }
                ]
            });
        var outputPath = CreateOutputPath("product-update.md");

        Assert.Equal(
            "Mention the pilot customers only after slide 2.",
            GetSpeakerNotesText(path, 0));
        Assert.Equal(
            "Do not read the backup metric in notes.",
            GetSpeakerNotesText(path, 1));

        var talkingPoints = await ExtractTalkingPointsAsync(path, topN: 3);
        var markdown = await _tools.pptx_export_markdown(path, outputPath);

        Assert.Equal(4, talkingPoints.Count);
        Assert.Equal(
            [
                "Launch completed in 3 regions",
                "Error rate dropped below 1%",
                "Support volume stayed flat"
            ],
            talkingPoints[0].Points);
        Assert.Equal(
            [
                "Enterprise expansion in APAC",
                "2 lighthouse customers went live",
                "Renewal intent reached 94%"
            ],
            talkingPoints[1].Points);
        Assert.Contains("Thin MCP tools delegate to PresentationService", talkingPoints[2].Points);
        Assert.Contains("OpenXML parsing stays centralized", talkingPoints[2].Points);

        Assert.Contains("# FY26 Launch Review", markdown);
        Assert.Contains("## Slide 2: Adoption Highlights", markdown);
        Assert.Contains("### Executive summary", markdown);
        Assert.Contains("- Launch completed in 3 regions", markdown);
        Assert.Contains("| Segment | Status |", markdown);
        Assert.Contains("| Enterprise | Growing |", markdown);
        Assert.Contains("![Picture", markdown);
        Assert.DoesNotContain("pilot customers only after slide 2", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("backup metric in notes", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("internal service names", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        Assert.Equal(markdown, File.ReadAllText(outputPath));
    }

    [Fact]
    public async Task Phase1Tools_VisualEdgeCaseDeck_HandlesEmptyAndImageOnlySlides()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                SpeakerNotesText = "This slide stays blank on purpose."
            },
            new TestSlideDefinition
            {
                SpeakerNotesText = "Narrate the system diagram live.",
                IncludeImage = true
            },
            new TestSlideDefinition
            {
                TitleText = "Appendix"
            });
        var outputPath = CreateOutputPath("visual-edge-cases.md");

        Assert.Equal("This slide stays blank on purpose.", GetSpeakerNotesText(path, 0));
        Assert.Equal("Narrate the system diagram live.", GetSpeakerNotesText(path, 1));

        var talkingPoints = await ExtractTalkingPointsAsync(path);
        var markdown = await _tools.pptx_export_markdown(path, outputPath);

        Assert.Equal(3, talkingPoints.Count);
        Assert.Empty(talkingPoints[0].Points);
        Assert.Empty(talkingPoints[1].Points);
        Assert.Equal(["Appendix"], talkingPoints[2].Points);

        Assert.Contains("## Slide 1: Untitled Slide 1", markdown);
        Assert.Contains("## Slide 2: Untitled Slide 2", markdown);
        Assert.Contains("## Slide 3: Appendix", markdown);
        Assert.Contains("![Picture", markdown);
        Assert.DoesNotContain("blank on purpose", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("system diagram live", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task Phase1Tools_UnicodeDeck_PreservesUnicodeContent()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "全球发布 🌍",
                SpeakerNotesText = "Do not say the private codename.",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.SubTitle,
                        Paragraphs = ["地域別アップデート"]
                    },
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "日本市場での利用率は42％", IsBullet = true },
                            new TestParagraphDefinition { Text = "Café teams requested résumé export", IsBullet = true },
                            new TestParagraphDefinition { Text = "Emoji reactions increased to 128 👍", IsBullet = true }
                        ]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Localized Examples",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "“Smart quotes” should stay intact", IsBullet = true },
                            new TestParagraphDefinition { Text = "naïve Bayes summaries remain opt-in", IsBullet = true },
                            new TestParagraphDefinition { Text = "São Paulo onboarding finished early", IsBullet = true }
                        ]
                    }
                ]
            });
        var outputPath = CreateOutputPath("unicode.md");

        Assert.Equal("Do not say the private codename.", GetSpeakerNotesText(path, 0));

        var talkingPoints = await ExtractTalkingPointsAsync(path, topN: 3);
        var markdown = await _tools.pptx_export_markdown(path, outputPath);

        Assert.Equal(2, talkingPoints.Count);
        Assert.Equal(
            [
                "日本市場での利用率は42％",
                "Café teams requested résumé export",
                "Emoji reactions increased to 128 👍"
            ],
            talkingPoints[0].Points);
        Assert.Equal(
            [
                "“Smart quotes” should stay intact",
                "naïve Bayes summaries remain opt-in",
                "São Paulo onboarding finished early"
            ],
            talkingPoints[1].Points);

        Assert.Contains("# 全球发布 🌍", markdown);
        Assert.Contains("### 地域別アップデート", markdown);
        Assert.Contains("- 日本市場での利用率は42％", markdown);
        Assert.Contains("- Café teams requested résumé export", markdown);
        Assert.Contains("- Emoji reactions increased to 128 👍", markdown);
        Assert.Contains("- “Smart quotes” should stay intact", markdown);
        Assert.Contains("- naïve Bayes summaries remain opt-in", markdown);
        Assert.Contains("- São Paulo onboarding finished early", markdown);
        Assert.DoesNotContain("private codename", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    private async Task<List<SlideTalkingPoints>> ExtractTalkingPointsAsync(string path, int topN = 5)
    {
        var result = await _tools.pptx_extract_talking_points(path, topN);
        var talkingPoints = JsonSerializer.Deserialize<List<SlideTalkingPoints>>(result);
        Assert.NotNull(talkingPoints);
        return talkingPoints;
    }

    private string CreateOutputPath(string fileName)
    {
        var directory = Path.Join(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(directory);
        TrackTempFile(directory);
        return Path.Join(directory, fileName);
    }

    private static string GetSpeakerNotesText(string path, int slideIndex)
    {
        using var document = PresentationDocument.Open(path, false);
        var presentationPart = document.PresentationPart!;
        var presentation = presentationPart.Presentation;
        Assert.NotNull(presentation);

        var slideIdList = presentation!.SlideIdList;
        Assert.NotNull(slideIdList);

        var slideId = slideIdList!.Elements<SlideId>().ElementAt(slideIndex);
        var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!);
        var notesSlidePart = slidePart.NotesSlidePart;
        Assert.NotNull(notesSlidePart);

        var notesSlide = notesSlidePart!.NotesSlide;
        Assert.NotNull(notesSlide);
        Assert.NotNull(notesSlide.CommonSlideData?.ShapeTree);

        var notesShape = notesSlide.CommonSlideData!.ShapeTree!.Elements<Shape>()
            .Single(shape =>
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == PlaceholderValues.Body);

        Assert.NotNull(notesShape.TextBody);
        return notesShape.TextBody!.InnerText;
    }
}
