using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxMcp.Tests.Services;

[Trait("Category", "Unit")]
public class UpdateSlideDataTests : PptxTestBase
{

    [Fact]
    public void UpdateSlideData_UpdatesMetricSlideByShapeName_AndKeepsPresentationCompatible()
    {
        var path = CreateRealMetricDeck();

        var result = Service.UpdateSlideData(path, 2, "Revenue Value", null, "4.8M ARR");

        Assert.True(result.Success);
        Assert.Equal(2, result.SlideNumber);
        Assert.Equal("shapeName", result.MatchedBy);
        Assert.Equal("Revenue Value", result.ResolvedShapeName);
        Assert.Equal("3.2M ARR", result.PreviousText);
        Assert.NotNull(result.ResolvedShapeId);
        Assert.Equal("4.8M ARR", FindShape(path, 1, "Revenue Value").Text);
        Assert.Equal("62%", FindShape(path, 1, "Gross Margin").Text);
        AssertPresentationCompatible(path);
    }

    [Fact]
    public void UpdateSlideData_UpdatesPlaceholderIndexesAcrossTitleContentAndDashboardStructures()
    {
        var path = CreateRealMetricDeck();

        var titleSlideResult = Service.UpdateSlideData(path, 1, null, 1, "Executive metrics review");
        var contentSlideResult = Service.UpdateSlideData(path, 3, null, 1, "Mitigate EMEA churn\nFinish finance automation");
        var dashboardSlideResult = Service.UpdateSlideData(path, 2, null, 4, "Pipeline conversion stable");

        Assert.True(titleSlideResult.Success);
        Assert.Equal("placeholderIndex", titleSlideResult.MatchedBy);
        Assert.Equal("Executive Subtitle", titleSlideResult.ResolvedShapeName);
        Assert.NotNull(titleSlideResult.PlaceholderType);
        Assert.Equal("Executive metrics review", FindShape(path, 0, "Executive Subtitle").Text);

        Assert.True(contentSlideResult.Success);
        Assert.Equal("placeholderIndex", contentSlideResult.MatchedBy);
        Assert.Equal("Risk Body", contentSlideResult.ResolvedShapeName);
        Assert.NotNull(contentSlideResult.PlaceholderType);
        Assert.Equal(
            ["Mitigate EMEA churn", "Finish finance automation"],
            FindShape(path, 2, "Risk Body").Paragraphs);

        Assert.True(dashboardSlideResult.Success);
        Assert.Equal("placeholderIndex", dashboardSlideResult.MatchedBy);
        Assert.Equal("Trend Commentary", dashboardSlideResult.ResolvedShapeName);
        Assert.NotNull(dashboardSlideResult.PlaceholderType);
        Assert.Equal("Pipeline conversion stable", FindShape(path, 1, "Trend Commentary").Text);

        AssertPresentationCompatible(path);
    }

    [Fact]
    public void UpdateSlideData_AppliesMultipleUpdatesToSameSlide_AndPreservesUntouchedShapes()
    {
        var path = CreateRealMetricDeck();
        var untouchedShapeBefore = GetTextBodyOuterXml(path, 1, "Trend Commentary");

        var firstResult = Service.UpdateSlideData(path, 2, "Revenue Value", null, "North America ARR up 18%");
        var secondResult = Service.UpdateSlideData(path, 2, "Missing metric", 3, "67%");
        var thirdResult = Service.UpdateSlideData(path, 2, "Owner Status", null, string.Empty);

        Assert.True(firstResult.Success);
        Assert.True(secondResult.Success);
        Assert.True(thirdResult.Success);
        Assert.Equal("shapeName", firstResult.MatchedBy);
        Assert.Equal("placeholderIndexFallback", secondResult.MatchedBy);
        Assert.Equal("shapeName", thirdResult.MatchedBy);

        Assert.Equal("North America ARR up 18%", FindShape(path, 1, "Revenue Value").Text);
        Assert.Equal("67%", FindShape(path, 1, "Gross Margin").Text);
        Assert.Equal(string.Empty, FindShape(path, 1, "Owner Status").Text);
        Assert.Equal(untouchedShapeBefore, GetTextBodyOuterXml(path, 1, "Trend Commentary"));
        Assert.Equal(1, CountPictures(path, 1));
        AssertPresentationCompatible(path);
    }

    [Fact]
    public void UpdateSlideData_ReturnsFailureForMissingShape_WithoutChangingSlideContents()
    {
        var path = CreateRealMetricDeck();
        var slideXmlBefore = Service.GetSlideXml(path, 1);

        var result = Service.UpdateSlideData(path, 2, "Missing Shape", null, "4.0M ARR");

        Assert.False(result.Success);
        Assert.Contains("Available shapes", result.Message);
        Assert.Contains("Revenue Value", result.Message);
        Assert.Equal("4.0M ARR", result.NewText);
        Assert.Equal(slideXmlBefore, Service.GetSlideXml(path, 1));
    }

    [Fact]
    public void UpdateSlideData_PreservesUnicodeEmojiAndLiteralHtmlEntityText()
    {
        var path = CreateRealMetricDeck();
        const string newText = "AT&amp;T < FY27 > 🚀 — café résumé";

        var result = Service.UpdateSlideData(path, 2, "Trend Commentary", null, newText);

        Assert.True(result.Success);
        Assert.Equal(newText, FindShape(path, 1, "Trend Commentary").Text);
        Assert.Equal(newText, GetShapeParagraphText(path, 1, "Trend Commentary"));
        AssertPresentationCompatible(path);
    }

    [Fact]
    public void UpdateSlideData_ClearsText_WhenEmptyStringIsRequested()
    {
        var path = CreateRealMetricDeck();

        var result = Service.UpdateSlideData(path, 2, "Owner Status", null, string.Empty);

        Assert.True(result.Success);
        Assert.Equal("Amber", result.PreviousText);
        Assert.Equal(string.Empty, FindShape(path, 1, "Owner Status").Text);
        Assert.Equal(string.Empty, GetShapeParagraphText(path, 1, "Owner Status"));
        Assert.Single(GetShapeParagraphs(path, 1, "Owner Status"));
        AssertPresentationCompatible(path);
    }

    [Fact]
    public void UpdateSlideData_CompletesSingleMetricUpdateUnder500Milliseconds()
    {
        var warmupPath = CreateRealMetricDeck();
        var measuredPath = CreateRealMetricDeck();

        Service.UpdateSlideData(warmupPath, 2, "Revenue Value", null, "4.0M ARR");

        var stopwatch = Stopwatch.StartNew();
        var result = Service.UpdateSlideData(measuredPath, 2, "Revenue Value", null, "4.1M ARR");
        stopwatch.Stop();

        Assert.True(result.Success);
        Assert.True(
            stopwatch.Elapsed.TotalMilliseconds < 500,
            $"Expected a single metric update to finish in under 500ms, but it took {stopwatch.Elapsed.TotalMilliseconds:F2}ms.");
        AssertPresentationCompatible(measuredPath);
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
                        Name = "Revenue Label",
                        Paragraphs = ["Net Revenue"],
                        X = Emu.OneInch,
                        Y = Emu.Inches1_5,
                        Width = Emu.Inches3,
                        Height = Emu.HalfInch
                    },
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
                        X = Emu.OneInch,
                        Y = Emu.Inches2,
                        Width = Emu.Inches3,
                        Height = Emu.ThreeQuartersInch
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Gross Margin",
                        Paragraphs = ["62%"],
                        X = Emu.Inches4,
                        Y = Emu.Inches2,
                        Width = Emu.Inches2,
                        Height = Emu.ThreeQuartersInch
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
                        X = Emu.OneInch,
                        Y = Emu.Inches3,
                        Width = Emu.Inches4,
                        Height = Emu.Inches1_25
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Owner Status",
                        Paragraphs = ["Amber"],
                        X = Emu.Inches5_5,
                        Y = Emu.Inches2,
                        Width = Emu.Inches1_5,
                        Height = Emu.ThreeQuartersInch
                    }
                ],
                Tables =
                [
                    new TestTableDefinition
                    {
                        Name = "Regional Breakdown",
                        Rows =
                        [
                            ["Region", "ARR"],
                            ["NA", "3.2M"],
                            ["EMEA", "1.4M"]
                        ],
                        X = Emu.Inches5,
                        Y = Emu.Inches3,
                        Width = Emu.Inches4,
                        Height = Emu.Inches1_5
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

    private static string GetTextBodyOuterXml(string path, int slideIndex, string shapeName)
    {
        using var document = PresentationDocument.Open(path, false);
        var shape = GetShapeElement(document, slideIndex, shapeName);
        return Assert.IsType<DocumentFormat.OpenXml.Presentation.TextBody>(shape.TextBody).OuterXml;
    }

    private static string GetShapeParagraphText(string path, int slideIndex, string shapeName)
    {
        using var document = PresentationDocument.Open(path, false);
        var shape = GetShapeElement(document, slideIndex, shapeName);
        return string.Join("\n", shape.TextBody!.Elements<A.Paragraph>().Select(paragraph => paragraph.InnerText));
    }

    private static IReadOnlyList<A.Paragraph> GetShapeParagraphs(string path, int slideIndex, string shapeName)
    {
        using var document = PresentationDocument.Open(path, false);
        var shape = GetShapeElement(document, slideIndex, shapeName);
        return shape.TextBody!.Elements<A.Paragraph>().ToList();
    }

    private static int CountPictures(string path, int slideIndex)
    {
        using var document = PresentationDocument.Open(path, false);
        var slidePart = GetSlidePart(document, slideIndex);
        var slide = Assert.IsType<Slide>(slidePart.Slide);
        var shapeTree = Assert.IsType<ShapeTree>(slide.CommonSlideData!.ShapeTree);
        return shapeTree.Elements<Picture>().Count();
    }

    private static Shape GetShapeElement(PresentationDocument document, int slideIndex, string shapeName)
    {
        var slidePart = GetSlidePart(document, slideIndex);
        var slide = Assert.IsType<Slide>(slidePart.Slide);
        var shapeTree = Assert.IsType<ShapeTree>(slide.CommonSlideData!.ShapeTree);
        return shapeTree.Elements<Shape>()
            .Single(shape => shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value == shapeName);
    }

    private static SlidePart GetSlidePart(PresentationDocument document, int slideIndex)
    {
        var presentationPart = Assert.IsType<PresentationPart>(document.PresentationPart);
        var presentation = Assert.IsType<Presentation>(presentationPart.Presentation);
        var slideIdList = Assert.IsType<SlideIdList>(presentation.SlideIdList);
        var slideId = slideIdList.Elements<SlideId>().ElementAt(slideIndex);
        return Assert.IsType<SlidePart>(presentationPart.GetPartById(slideId.RelationshipId!.Value!));
    }
}
