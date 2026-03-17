using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using PptxMcp.Models;

namespace PptxMcp.Tests.Services;

public class BatchUpdateTests : PptxTestBase
{

    [Fact]
    public void BatchUpdate_UpdatesMultipleSlides_AndKeepsPresentationCompatible()
    {
        var path = CreateRealMetricDeck();
        var baselineValidationErrors = ValidatePresentation(path);

        var result = Service.BatchUpdate(path,
        [
            new BatchUpdateMutation(1, "Executive Subtitle", "Executive metrics review"),
            new BatchUpdateMutation(2, "Revenue Value", "4.8M ARR"),
            new BatchUpdateMutation(3, "Risk Body", "Mitigate EMEA churn\nFinish finance automation")
        ]);

        Assert.Equal(3, result.TotalMutations);
        Assert.Equal(3, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
        Assert.Collection(result.Results,
            mutationResult =>
            {
                Assert.Equal(1, mutationResult.SlideNumber);
                Assert.Equal("Executive Subtitle", mutationResult.ShapeName);
                Assert.True(mutationResult.Success);
                Assert.Null(mutationResult.Error);
                Assert.Equal("shapeName", mutationResult.MatchedBy);
            },
            mutationResult =>
            {
                Assert.Equal(2, mutationResult.SlideNumber);
                Assert.Equal("Revenue Value", mutationResult.ShapeName);
                Assert.True(mutationResult.Success);
                Assert.Null(mutationResult.Error);
                Assert.Equal("shapeName", mutationResult.MatchedBy);
            },
            mutationResult =>
            {
                Assert.Equal(3, mutationResult.SlideNumber);
                Assert.Equal("Risk Body", mutationResult.ShapeName);
                Assert.True(mutationResult.Success);
                Assert.Null(mutationResult.Error);
                Assert.Equal("shapeName", mutationResult.MatchedBy);
            });

        Assert.Equal("Executive metrics review", FindShape(path, 0, "Executive Subtitle").Text);
        Assert.Equal("4.8M ARR", FindShape(path, 1, "Revenue Value").Text);
        Assert.Equal(
            ["Mitigate EMEA churn", "Finish finance automation"],
            FindShape(path, 2, "Risk Body").Paragraphs);
        AssertPresentationCompatible(path);
        Assert.Equal(baselineValidationErrors, ValidatePresentation(path));
    }

    [Fact]
    public void BatchUpdate_PreservesSuccessfulMutations_WhenOneMutationFails()
    {
        var path = CreateRealMetricDeck();
        var failedSlideXmlBefore = Service.GetSlideXml(path, 1);

        var result = Service.BatchUpdate(path,
        [
            new BatchUpdateMutation(1, "Executive Subtitle", "Executive metrics review"),
            new BatchUpdateMutation(2, "Missing Shape", "4.0M ARR"),
            new BatchUpdateMutation(3, "Risk Body", "Mitigate EMEA churn")
        ]);

        Assert.Equal(3, result.TotalMutations);
        Assert.Equal(2, result.SuccessCount);
        Assert.Equal(1, result.FailureCount);
        Assert.Collection(result.Results,
            mutationResult =>
            {
                Assert.True(mutationResult.Success);
                Assert.Null(mutationResult.Error);
                Assert.Equal("shapeName", mutationResult.MatchedBy);
            },
            mutationResult =>
            {
                Assert.Equal(2, mutationResult.SlideNumber);
                Assert.Equal("Missing Shape", mutationResult.ShapeName);
                Assert.False(mutationResult.Success);
                Assert.Null(mutationResult.MatchedBy);
                Assert.Contains("Available shapes", mutationResult.Error);
            },
            mutationResult =>
            {
                Assert.True(mutationResult.Success);
                Assert.Null(mutationResult.Error);
                Assert.Equal("shapeName", mutationResult.MatchedBy);
            });

        Assert.Equal("Executive metrics review", FindShape(path, 0, "Executive Subtitle").Text);
        Assert.Equal("3.2M ARR", FindShape(path, 1, "Revenue Value").Text);
        Assert.Equal("62%", FindShape(path, 1, "Gross Margin").Text);
        Assert.Equal(["Mitigate EMEA churn"], FindShape(path, 2, "Risk Body").Paragraphs);
        Assert.Equal(failedSlideXmlBefore, Service.GetSlideXml(path, 1));
        AssertPresentationCompatible(path);
    }

    [Fact]
    public void BatchUpdate_ReturnsZeroCounts_ForEmptyBatch()
    {
        var path = CreateRealMetricDeck();
        var baselineValidationErrors = ValidatePresentation(path);

        var result = Service.BatchUpdate(path, []);

        Assert.Equal(0, result.TotalMutations);
        Assert.Equal(0, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
        Assert.Empty(result.Results);
        Assert.Equal("Board dashboard", FindShape(path, 0, "Executive Subtitle").Text);
        Assert.Equal("3.2M ARR", FindShape(path, 1, "Revenue Value").Text);
        AssertPresentationCompatible(path);
        Assert.Equal(baselineValidationErrors, ValidatePresentation(path));
    }

    private string CreateRealMetricDeck() =>
        CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "FY26 Metrics",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Executive Subtitle",
                        PlaceholderType = PlaceholderValues.SubTitle,
                        Paragraphs = ["Board dashboard"]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Revenue Dashboard",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Revenue Value",
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition
                            {
                                Text = "3.2M ARR",
                                IsBullet = true,
                                Level = 1
                            }
                        ],
                        X = 914400,
                        Y = 1828800,
                        Width = 2743200,
                        Height = 685800
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Gross Margin",
                        Paragraphs = ["62%"],
                        X = 3657600,
                        Y = 1828800,
                        Width = 1828800,
                        Height = 685800
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Trend Commentary",
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Pipeline stable", IsBullet = true },
                            new TestParagraphDefinition { Text = "Upsell motion healthy", IsBullet = true }
                        ],
                        X = 914400,
                        Y = 2743200,
                        Width = 3657600,
                        Height = 1143000
                    }
                ],
                IncludeImage = true
            },
            new TestSlideDefinition
            {
                TitleText = "Execution Risks",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Risk Body",
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition { Text = "Support EMEA renewals", IsBullet = true },
                            new TestParagraphDefinition { Text = "Close finance automation gap", IsBullet = true }
                        ]
                    }
                ]
            });

    private ShapeContent FindShape(string path, int slideIndex, string shapeName)
    {
        var slide = Service.GetSlideContent(path, slideIndex);
        return Assert.Single(slide.Shapes, shape => shape.Name == shapeName);
    }

    private static void AssertPresentationCompatible(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var presentationPart = Assert.IsType<PresentationPart>(document.PresentationPart);
        var presentation = Assert.IsType<Presentation>(presentationPart.Presentation);
        var slideIdList = Assert.IsType<SlideIdList>(presentation.SlideIdList);
        var slideIds = slideIdList.Elements<SlideId>().ToList();
        Assert.Equal(3, slideIds.Count);

        foreach (var slideId in slideIds)
        {
            var slidePart = Assert.IsType<SlidePart>(presentationPart.GetPartById(slideId.RelationshipId!.Value!));
            var slide = Assert.IsType<Slide>(slidePart.Slide);
            Assert.NotNull(slide.CommonSlideData?.ShapeTree);
        }
    }

    private static List<string> ValidatePresentation(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var validator = new OpenXmlValidator();
        return validator.Validate(document)
            .Select(error => $"{error.Path?.XPath ?? "<unknown>"}: {error.Description}")
            .ToList();
    }
}
