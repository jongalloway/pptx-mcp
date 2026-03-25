using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;

namespace PptxTools.Tests.Tools;

[Trait("Category", "E2E")]
public class PptxPhase2E2eTests : PptxTestBase
{
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public PptxPhase2E2eTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    [Fact]
    public async Task Phase2Workflow_MultiSourceDeck_ReadsFetchesUpdatesAndVerifiesAcrossMultipleSlides()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Q3 Operating Review",
                SpeakerNotesText = "Keep the board-only forecast in the appendix.",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.SubTitle,
                        Paragraphs = ["Prepared for the weekly exec sync"]
                    },
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Revenue trend lagged forecast in July", IsBullet = true },
                            new TestParagraphDefinition { Text = "NPS follow-up completed with top accounts", IsBullet = true },
                            new TestParagraphDefinition { Text = "Staffing plan remained stable through August", IsBullet = true }
                        ]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "KPI Dashboard",
                SpeakerNotesText = "CFO asked for a private margin bridge.",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Revenue Value",
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Revenue growth: 12%", IsBullet = true }
                        ]
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Nps Value",
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Net promoter score: 54", IsBullet = true }
                        ]
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Pipeline Value",
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Pipeline coverage: 3.1x", IsBullet = true }
                        ]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Team Updates",
                SpeakerNotesText = "Mention the headcount freeze only if asked.",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Launch Status",
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Launch status: Yellow — pilot exit criteria open", IsBullet = true }
                        ]
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Hiring Status",
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Hiring status: On plan for 2 backend roles", IsBullet = true }
                        ]
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Customer Notes",
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Customer asks: analytics export and SSO", IsBullet = true }
                        ]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "What's New",
                SpeakerNotesText = "Skip the embargoed roadmap item.",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "What's New Body",
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Older .NET preview highlights", IsBullet = true },
                            new TestParagraphDefinition { Text = "Prior MCP SDK docs refresh", IsBullet = true }
                        ]
                    }
                ]
            });
        var outputPath = CreateOutputPath("phase2-multi-source.md");

        var initialTalkingPoints = await ExtractTalkingPointsAsync(path, topN: 3);
        var metricsSlideBefore = await GetSlideContentAsync(path, 1);
        var teamSlideBefore = await GetSlideContentAsync(path, 2);
        var whatsNewSlideBefore = await GetSlideContentAsync(path, 3);
        var initialMarkdown = await _tools.pptx_export_markdown(path, outputPath);
        var baselineValidationErrors = ValidatePresentation(path);

        Assert.Equal(4, initialTalkingPoints.Count);
        Assert.Contains("Revenue growth: 12%", initialTalkingPoints[1].Points);
        Assert.Contains("Launch status: Yellow — pilot exit criteria open", initialTalkingPoints[2].Points);
        Assert.Equal("Revenue growth: 12%", FindShape(metricsSlideBefore, "Revenue Value").Text);
        Assert.Equal("Hiring status: On plan for 2 backend roles", FindShape(teamSlideBefore, "Hiring Status").Text);
        Assert.Equal(
            ["Older .NET preview highlights", "Prior MCP SDK docs refresh"],
            FindShape(whatsNewSlideBefore, "What's New Body").Paragraphs);
        Assert.Contains("# Q3 Operating Review", initialMarkdown);
        Assert.Contains("Revenue growth: 12%", initialMarkdown);
        Assert.Contains("Older .NET preview highlights", initialMarkdown);
        Assert.DoesNotContain("private margin bridge", initialMarkdown, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("headcount freeze", initialMarkdown, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("embargoed roadmap item", initialMarkdown, StringComparison.OrdinalIgnoreCase);

        var metricsSnapshot = new MockMetricSnapshot(
            RevenueGrowth: "18%",
            NetPromoterScore: "67",
            PipelineCoverage: "4.3x");
        var teamSnapshot = new MockTeamSnapshot(
            LaunchStatus: "Green — pilot expanded to 3 regions",
            HiringStatus: "Backfill approved for solutions architect",
            CustomerRequests: "SSO export, audit history, PowerPoint diffing");
        MockPost[] latestPosts =
        [
            new MockPost(
                "Introducing MCP Composition Patterns for .NET Agents",
                "Learn how to compose multiple MCP servers to build powerful multi-source agent workflows."),
            new MockPost(
                "What's New in .NET 10 Preview 4",
                "Performance improvements in JIT and GC, new LINQ overloads, and System.AI namespace."),
            new MockPost(
                "Building Intelligent Agents with ModelContextProtocol SDK",
                "Deep dive into the MCP SDK: tool registration, stdio transport, and composable tools.")
        ];

        var revenueUpdate = await UpdateSlideDataAsync(path, 2, "Revenue Value", null, $"Revenue growth: {metricsSnapshot.RevenueGrowth}");
        var npsUpdate = await UpdateSlideDataAsync(path, 2, "Nps Value", null, $"Net promoter score: {metricsSnapshot.NetPromoterScore}");
        var pipelineUpdate = await UpdateSlideDataAsync(path, 2, "Pipeline Value", null, $"Pipeline coverage: {metricsSnapshot.PipelineCoverage}");
        var launchUpdate = await UpdateSlideDataAsync(path, 3, "Launch Status", null, $"Launch status: {teamSnapshot.LaunchStatus}");
        var hiringUpdate = await UpdateSlideDataAsync(path, 3, "Hiring Status", null, $"Hiring status: {teamSnapshot.HiringStatus}");
        var customerNotesUpdate = await UpdateSlideDataAsync(path, 3, "Customer Notes", null, $"Customer asks: {teamSnapshot.CustomerRequests}");
        var whatsNewBodyIndex = FindTextShapeIndex(whatsNewSlideBefore, "What's New Body");
        var whatsNewUpdate = await UpdateSlideDataAsync(
            path,
            4,
            "Missing What's New Body",
            whatsNewBodyIndex,
            string.Join(Environment.NewLine, latestPosts.Select(post => $"{post.Title} — {post.Summary}")));

        Assert.True(revenueUpdate.Success);
        Assert.Equal("shapeName", revenueUpdate.MatchedBy);
        Assert.Equal("Revenue growth: 12%", revenueUpdate.PreviousText);
        Assert.True(npsUpdate.Success);
        Assert.True(pipelineUpdate.Success);
        Assert.True(launchUpdate.Success);
        Assert.True(hiringUpdate.Success);
        Assert.True(customerNotesUpdate.Success);
        Assert.True(whatsNewUpdate.Success);
        Assert.Equal("placeholderIndexFallback", whatsNewUpdate.MatchedBy);
        Assert.Equal("What's New Body", whatsNewUpdate.ResolvedShapeName);
        Assert.Equal("Older .NET preview highlights\nPrior MCP SDK docs refresh", whatsNewUpdate.PreviousText);

        var metricsSlideAfter = await GetSlideContentAsync(path, 1);
        var teamSlideAfter = await GetSlideContentAsync(path, 2);
        var whatsNewSlideAfter = await GetSlideContentAsync(path, 3);
        var talkingPointsAfter = await ExtractTalkingPointsAsync(path, topN: 3);
        var markdownAfter = await _tools.pptx_export_markdown(path, outputPath);

        Assert.Equal("Revenue growth: 18%", FindShape(metricsSlideAfter, "Revenue Value").Text);
        Assert.Equal("Net promoter score: 67", FindShape(metricsSlideAfter, "Nps Value").Text);
        Assert.Equal("Pipeline coverage: 4.3x", FindShape(metricsSlideAfter, "Pipeline Value").Text);
        Assert.Equal("Launch status: Green — pilot expanded to 3 regions", FindShape(teamSlideAfter, "Launch Status").Text);
        Assert.Equal("Hiring status: Backfill approved for solutions architect", FindShape(teamSlideAfter, "Hiring Status").Text);
        Assert.Equal("Customer asks: SSO export, audit history, PowerPoint diffing", FindShape(teamSlideAfter, "Customer Notes").Text);
        Assert.Equal(
            latestPosts.Select(post => $"{post.Title} — {post.Summary}").ToList(),
            FindShape(whatsNewSlideAfter, "What's New Body").Paragraphs);

        Assert.Contains("Revenue growth: 18%", talkingPointsAfter[1].Points);
        Assert.Contains("Net promoter score: 67", talkingPointsAfter[1].Points);
        Assert.Contains("Launch status: Green — pilot expanded to 3 regions", talkingPointsAfter[2].Points);
        Assert.Contains(latestPosts[0].Title, talkingPointsAfter[3].Points[0]);

        Assert.Contains("Revenue growth: 18%", markdownAfter);
        Assert.Contains("Net promoter score: 67", markdownAfter);
        Assert.Contains("Pipeline coverage: 4.3x", markdownAfter);
        Assert.Contains("Hiring status: Backfill approved for solutions architect", markdownAfter);
        Assert.Contains("Customer asks: SSO export, audit history, PowerPoint diffing", markdownAfter);
        Assert.Contains(latestPosts[0].Title, markdownAfter);
        Assert.Contains(latestPosts[1].Title, markdownAfter);
        Assert.Contains(latestPosts[2].Title, markdownAfter);
        Assert.DoesNotContain("Revenue growth: 12%", markdownAfter);
        Assert.DoesNotContain("Net promoter score: 54", markdownAfter);
        Assert.DoesNotContain("Pipeline coverage: 3.1x", markdownAfter);
        Assert.DoesNotContain("Launch status: Yellow — pilot exit criteria open", markdownAfter);
        Assert.DoesNotContain("Older .NET preview highlights", markdownAfter);
        Assert.DoesNotContain("private margin bridge", markdownAfter, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("headcount freeze", markdownAfter, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("embargoed roadmap item", markdownAfter, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(markdownAfter, File.ReadAllText(outputPath));

        Assert.Equal("CFO asked for a private margin bridge.", GetSpeakerNotesText(path, 1));
        Assert.Equal("Mention the headcount freeze only if asked.", GetSpeakerNotesText(path, 2));
        Assert.Equal("Skip the embargoed roadmap item.", GetSpeakerNotesText(path, 3));

        AssertPresentationCanBeOpened(path, expectedSlideCount: 4);
        Assert.Equal(baselineValidationErrors, ValidatePresentation(path));
    }

    private async Task<SlideContent> GetSlideContentAsync(string path, int slideIndex)
    {
        var result = await _tools.pptx_get_slide_content(path, slideIndex);
        var slideContent = JsonSerializer.Deserialize<SlideContent>(result);
        Assert.NotNull(slideContent);
        return slideContent;
    }

    private async Task<SlideDataUpdateResult> UpdateSlideDataAsync(string path, int slideNumber, string? shapeName, int? placeholderIndex, string newText)
    {
        var result = await _tools.pptx_update_slide_data(path, slideNumber, shapeName, placeholderIndex, newText);
        var updateResult = JsonSerializer.Deserialize<SlideDataUpdateResult>(result);
        Assert.NotNull(updateResult);
        return updateResult;
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

    private static ShapeContent FindShape(SlideContent slideContent, string shapeName) =>
        Assert.Single(slideContent.Shapes, shape => string.Equals(shape.Name, shapeName, StringComparison.Ordinal));

    private static int FindTextShapeIndex(SlideContent slideContent, string shapeName)
    {
        var index = slideContent.Shapes
            .Select((shape, position) => new { shape, position })
            .Where(entry => entry.shape.Text is not null)
            .Single(entry => string.Equals(entry.shape.Name, shapeName, StringComparison.Ordinal))
            .position;

        return index;
    }

    private static void AssertPresentationCanBeOpened(string path, int expectedSlideCount)
    {
        using var document = PresentationDocument.Open(path, false);
        var presentationPart = document.PresentationPart;
        Assert.NotNull(presentationPart);
        var presentation = presentationPart!.Presentation;
        Assert.NotNull(presentation);
        var slideIds = presentation!.SlideIdList?.Elements<SlideId>().ToList();
        Assert.NotNull(slideIds);
        Assert.Equal(expectedSlideCount, slideIds!.Count);

        Assert.All(slideIds, slideId =>
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!);
            Assert.NotNull(slidePart.Slide);
            Assert.NotNull(slidePart.Slide.CommonSlideData?.ShapeTree);
        });
    }

    private static List<string> ValidatePresentation(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var validator = new OpenXmlValidator();
        return validator.Validate(document)
            .Select(error => $"{error.Path?.XPath ?? "<unknown>"}: {error.Description}")
            .ToList();
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

    private sealed record MockMetricSnapshot(
        string RevenueGrowth,
        string NetPromoterScore,
        string PipelineCoverage);

    private sealed record MockTeamSnapshot(
        string LaunchStatus,
        string HiringStatus,
        string CustomerRequests);

    private sealed record MockPost(
        string Title,
        string Summary);
}
