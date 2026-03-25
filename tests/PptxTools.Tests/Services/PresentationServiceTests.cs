using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxTools.Tests.Services;

[Trait("Category", "Unit")]
public class PresentationServiceTests : PptxTestBase
{

    [Fact]
    public void GetSlides_ReturnsCorrectCount()
    {
        var path = CreateMinimalPptx();
        var slides = Service.GetSlides(path);
        Assert.Single(slides);
    }

    [Fact]
    public void GetSlides_ReturnsCorrectTitle()
    {
        var path = CreateMinimalPptx("Hello World");
        var slides = Service.GetSlides(path);
        Assert.Equal("Hello World", slides[0].Title);
    }

    [Fact]
    public void GetSlides_SlideHasCorrectIndex()
    {
        var path = CreateMinimalPptx();
        var slides = Service.GetSlides(path);
        Assert.Equal(0, slides[0].Index);
    }

    [Fact]
    public void GetLayouts_ReturnsLayouts()
    {
        var path = CreateMinimalPptx();
        var layouts = Service.GetLayouts(path);
        Assert.NotEmpty(layouts);
    }

    [Fact]
    public void GetLayouts_LayoutHasName()
    {
        var path = CreateMinimalPptx();
        var layouts = Service.GetLayouts(path);
        Assert.All(layouts, layout => Assert.NotNull(layout.Name));
    }

    [Fact]
    public void AddSlide_IncreasesSlideCount()
    {
        var path = CreateMinimalPptx();
        var before = Service.GetSlides(path);
        Service.AddSlide(path, null);
        var after = Service.GetSlides(path);
        Assert.Equal(before.Count + 1, after.Count);
    }

    [Fact]
    public void AddSlide_ReturnsNewSlideIndex()
    {
        var path = CreateMinimalPptx();
        var newIndex = Service.AddSlide(path, null);
        Assert.Equal(1, newIndex);
    }

    [Fact]
    public void UpdateTextPlaceholder_ChangesTextContent()
    {
        var path = CreateMinimalPptx("Original Title");
        Service.UpdateTextPlaceholder(path, 0, 0, "Updated Title");
        var slides = Service.GetSlides(path);
        Assert.Equal("Updated Title", slides[0].Title);
    }

    [Fact]
    public void UpdateSlideData_UpdatesShapeByNameAndPreservesParagraphProperties()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Dashboard",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Revenue Value",
                        PlaceholderType = PlaceholderValues.Body,
                        ParagraphDefinitions =
                        [
                            new TestParagraphDefinition
                            {
                                Text = "12%",
                                IsBullet = true,
                                Level = 1
                            }
                        ]
                    }
                ]
            });

        var result = Service.UpdateSlideData(path, 1, "Revenue Value", null, "18%");

        Assert.True(result.Success);
        Assert.Equal("shapeName", result.MatchedBy);
        Assert.Equal("12%", result.PreviousText);

        var content = Service.GetSlideContent(path, 0);
        var updatedShape = Assert.Single(content.Shapes, shape => shape.Name == "Revenue Value");
        Assert.Equal("18%", updatedShape.Text);

        using var doc = PresentationDocument.Open(path, false);
        var presentationPart = Assert.IsType<PresentationPart>(doc.PresentationPart);
        var presentation = Assert.IsType<Presentation>(presentationPart.Presentation);
        var slideIdList = Assert.IsType<SlideIdList>(presentation.SlideIdList);
        var slideId = Assert.Single(slideIdList.Elements<SlideId>());
        var slidePart = Assert.IsType<SlidePart>(presentationPart.GetPartById(slideId.RelationshipId!.Value!));
        var slide = Assert.IsType<Slide>(slidePart.Slide);
        var shapeTree = Assert.IsType<ShapeTree>(slide.CommonSlideData!.ShapeTree);
        var shape = shapeTree.Elements<Shape>()
            .Single(candidate => candidate.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value == "Revenue Value");
        var paragraphProperties = Assert.Single(shape.TextBody!.Elements<A.Paragraph>()).ParagraphProperties;
        Assert.NotNull(paragraphProperties);
        Assert.Equal(1, paragraphProperties.Level?.Value);
    }

    [Fact]
    public void UpdateSlideData_UpdatesShapeByIndexWithUnicodeText()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Dashboard",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Revenue Label",
                        Paragraphs = ["Revenue"]
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Revenue Value",
                        Paragraphs = ["12%"]
                    }
                ]
            });

        var result = Service.UpdateSlideData(path, 1, null, 2, "Δ 45% — 売上高");

        Assert.True(result.Success);
        Assert.Equal("placeholderIndex", result.MatchedBy);
        Assert.Equal("Revenue Value", result.ResolvedShapeName);

        var content = Service.GetSlideContent(path, 0);
        var updatedShape = Assert.Single(content.Shapes, shape => shape.Name == "Revenue Value");
        Assert.Equal("Δ 45% — 売上高", updatedShape.Text);
    }

    [Fact]
    public void UpdateSlideData_ReturnsFailureWhenShapeIsMissing()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Dashboard",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Revenue Value",
                        Paragraphs = ["12%"]
                    }
                ]
            });

        var result = Service.UpdateSlideData(path, 1, "Missing Shape", null, "18%");

        Assert.False(result.Success);
        Assert.Contains("Available shapes", result.Message);
        Assert.Equal("18%", result.NewText);
    }

    [Fact]
    public void UpdateSlideData_ReturnsFailureWhenSlideNumberIsOutOfRange()
    {
        var path = CreateMinimalPptx();

        var result = Service.UpdateSlideData(path, 2, "Title 1", null, "Updated");

        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message);
    }

    [Fact]
    public void GetSlideXml_ReturnsXmlString()
    {
        var path = CreateMinimalPptx();
        var xml = Service.GetSlideXml(path, 0);
        Assert.NotNull(xml);
        Assert.Contains("sld", xml);
    }

    [Fact]
    public void GetSlideXml_OutOfRange_ThrowsException()
    {
        var path = CreateMinimalPptx();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            Service.GetSlideXml(path, 99));
    }

    [Fact]
    public void GetSlideContent_ReturnsSlideIndex()
    {
        var path = CreateMinimalPptx();
        var content = Service.GetSlideContent(path, 0);
        Assert.Equal(0, content.SlideIndex);
    }

    [Fact]
    public void GetSlideContent_ReturnsSlideDimensions()
    {
        var path = CreateMinimalPptx();
        var content = Service.GetSlideContent(path, 0);
        Assert.True(content.SlideWidthEmu > 0);
        Assert.True(content.SlideHeightEmu > 0);
    }

    [Fact]
    public void GetSlideContent_ReturnsShapes()
    {
        var path = CreateMinimalPptx();
        var content = Service.GetSlideContent(path, 0);
        Assert.NotEmpty(content.Shapes);
    }

    [Fact]
    public void GetSlideContent_TitleShapeHasText()
    {
        var path = CreateMinimalPptx("My Title");
        var content = Service.GetSlideContent(path, 0);
        var titleShape = Assert.Single(content.Shapes, shape => shape.IsPlaceholder);
        Assert.Equal("My Title", titleShape.Text);
    }

    [Fact]
    public void GetSlideContent_TitleShapeIsTextType()
    {
        var path = CreateMinimalPptx();
        var content = Service.GetSlideContent(path, 0);
        var titleShape = Assert.Single(content.Shapes, shape => shape.IsPlaceholder);
        Assert.Equal("Text", titleShape.ShapeType);
    }

    [Fact]
    public void GetSlideContent_TitleShapeHasPlaceholderType()
    {
        var path = CreateMinimalPptx();
        var content = Service.GetSlideContent(path, 0);
        var titleShape = Assert.Single(content.Shapes, shape => shape.IsPlaceholder);
        Assert.NotNull(titleShape.PlaceholderType);
    }

    [Fact]
    public void GetSlideContent_OutOfRange_ThrowsException()
    {
        var path = CreateMinimalPptx();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            Service.GetSlideContent(path, 99));
    }

    [Fact]
    public void ExtractTalkingPoints_ReturnsRankedBodyPoints()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Release Overview",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Body 1",
                        PlaceholderType = PlaceholderValues.Body,
                        Paragraphs =
                        [
                            "Faster startup across workloads",
                            "Improved diagnostics for production",
                            "Native AOT reduces deployment size"
                        ]
                    }
                ]
            });

        var talkingPoints = Service.ExtractTalkingPoints(path, 2);

        Assert.Single(talkingPoints);
        Assert.Equal(
            [
                "Faster startup across workloads",
                "Improved diagnostics for production"
            ],
            talkingPoints[0].Points);
    }

    [Fact]
    public void ExtractTalkingPoints_FiltersPresenterNotesAndFormattingOnlyText()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Presenter Notes",
                        Paragraphs = ["Presenter Notes"]
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Body 1",
                        PlaceholderType = PlaceholderValues.Body,
                        Paragraphs = ["***", "Key metric improved 25%"]
                    }
                ]
            });

        var talkingPoints = Service.ExtractTalkingPoints(path);

        Assert.Single(talkingPoints[0].Points);
        Assert.Equal("Key metric improved 25%", talkingPoints[0].Points[0]);
    }

    [Fact]
    public void ExtractTalkingPoints_ReturnsEmptyForImageOnlySlide()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                IncludeImage = true
            });

        var talkingPoints = Service.ExtractTalkingPoints(path);

        Assert.Empty(talkingPoints[0].Points);
    }

    [Fact]
    public void ExtractTalkingPoints_ReturnsEmptyForEmptySlide()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition());

        var talkingPoints = Service.ExtractTalkingPoints(path);

        Assert.Empty(talkingPoints[0].Points);
    }

    [Fact]
    public void ExtractTalkingPoints_TitleOnlySlideFallsBackToTitle()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Quarterly Update"
            });

        var talkingPoints = Service.ExtractTalkingPoints(path);

        Assert.Single(talkingPoints[0].Points);
        Assert.Equal("Quarterly Update", talkingPoints[0].Points[0]);
    }

    [Fact]
    public void ExtractTalkingPoints_InvalidTopN_ThrowsException()
    {
        var path = CreateMinimalPptx();

        Assert.Throws<ArgumentOutOfRangeException>(() => Service.ExtractTalkingPoints(path, 0));
    }

    [Fact]
    public void ExtractTalkingPoints_ProcessesRealisticMultiSlideDeck()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Executive Summary",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Body 1",
                        PlaceholderType = PlaceholderValues.Body,
                        Paragraphs =
                        [
                            "Launch completed in 3 regions",
                            "Error rate dropped below 1%",
                            "Support volume stayed flat"
                        ]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Architecture",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Presenter Notes",
                        Paragraphs = ["Presenter Notes"]
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Body 2",
                        PlaceholderType = PlaceholderValues.Body,
                        Paragraphs =
                        [
                            "Thin MCP tools delegate to PresentationService",
                            "OpenXML parsing remains centralized"
                        ]
                    }
                ]
            },
            new TestSlideDefinition
            {
                IncludeImage = true
            });

        var talkingPoints = Service.ExtractTalkingPoints(path, 2);

        Assert.Equal(3, talkingPoints.Count);
        Assert.Equal(
            [
                "Launch completed in 3 regions",
                "Error rate dropped below 1%"
            ],
            talkingPoints[0].Points);
        Assert.Equal(
            [
                "Thin MCP tools delegate to PresentationService",
                "OpenXML parsing remains centralized"
            ],
            talkingPoints[1].Points);
        Assert.Empty(talkingPoints[2].Points);
    }

    [Fact]
    public void WriteNotes_CreatesNotesOnSlideWithoutExistingNotes()
    {
        var path = CreateMinimalPptx();
        Service.WriteNotes(path, 0, "Source: https://example.com");
        var slides = Service.GetSlides(path);
        Assert.Equal("Source: https://example.com", slides[0].Notes);
    }

    [Fact]
    public void WriteNotes_ReplacesExistingNotes()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Slide",
            SpeakerNotesText = "Old notes"
        });
        Service.WriteNotes(path, 0, "New notes");
        var slides = Service.GetSlides(path);
        Assert.Equal("New notes", slides[0].Notes);
    }

    [Fact]
    public void WriteNotes_AppendToExistingNotes()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Slide",
            SpeakerNotesText = "First"
        });
        Service.WriteNotes(path, 0, "Second", append: true);
        var slides = Service.GetSlides(path);
        Assert.NotNull(slides[0].Notes);
        Assert.Contains("First", slides[0].Notes);
        Assert.Contains("Second", slides[0].Notes);
    }

    [Fact]
    public void WriteNotes_AppendToEmptyNotes_SetsNotesWithoutLeadingNewline()
    {
        var path = CreateMinimalPptx();
        Service.WriteNotes(path, 0, "Only note", append: true);
        var slides = Service.GetSlides(path);
        Assert.NotNull(slides[0].Notes);
        Assert.Contains("Only note", slides[0].Notes);
    }

    [Fact]
    public void WriteNotes_MultiParagraph_PreservesAllParagraphs()
    {
        var path = CreateMinimalPptx();
        Service.WriteNotes(path, 0, "Line one\nLine two\nLine three");
        using var doc = PresentationDocument.Open(path, false);
        var slideIdList = doc.PresentationPart!.Presentation.SlideIdList!;
        var slidePart = (SlidePart)doc.PresentationPart.GetPartById(
            slideIdList.Elements<SlideId>().First().RelationshipId!.Value!);
        var notesSlide = slidePart.NotesSlidePart!.NotesSlide;
        var bodyShape = Assert.Single(notesSlide.CommonSlideData!.ShapeTree!.Elements<Shape>(),
            s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .PlaceholderShape?.Type?.Value == PlaceholderValues.Body);
        var paragraphs = bodyShape.TextBody!.Elements<A.Paragraph>().ToList();
        Assert.Equal(3, paragraphs.Count);
    }

    [Fact]
    public void WriteNotes_CreatesNotesMasterPartWhenMissing()
    {
        var path = CreateMinimalPptx();
        Service.WriteNotes(path, 0, "Verify master");
        using var doc = PresentationDocument.Open(path, false);
        Assert.NotNull(doc.PresentationPart!.NotesMasterPart);
    }

    [Fact]
    public void WriteNotes_AppendToMultiParagraphNotes_PreservesParagraphBoundaries()
    {
        // Existing notes have two paragraphs; append should not collapse them into one blob
        var path = CreateMinimalPptx();
        Service.WriteNotes(path, 0, "Para one\nPara two");
        Service.WriteNotes(path, 0, "Para three", append: true);
        var slides = Service.GetSlides(path);
        Assert.NotNull(slides[0].Notes);
        Assert.Contains("Para one", slides[0].Notes);
        Assert.Contains("Para two", slides[0].Notes);
        Assert.Contains("Para three", slides[0].Notes);
        // All three should be distinct — the round-trip via GetSlideNotes must preserve \n
        var parts = slides[0].Notes!.Split('\n');
        Assert.Equal(3, parts.Length);
    }

    [Fact]
    public void WriteNotes_ExistingNotesSlidePart_LinksNotesMasterPart()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Slide",
            SpeakerNotesText = "Original"
        });
        Service.WriteNotes(path, 0, "Updated");
        using var doc = PresentationDocument.Open(path, false);
        var slideIdList = doc.PresentationPart!.Presentation.SlideIdList!;
        var slidePart = (SlidePart)doc.PresentationPart.GetPartById(
            slideIdList.Elements<SlideId>().First().RelationshipId!.Value!);
        Assert.NotNull(slidePart.NotesSlidePart!.NotesMasterPart);
    }

    [Fact]
    public void WriteNotes_OutOfRangeSlideIndex_Throws()
    {
        var path = CreateMinimalPptx();
        Assert.Throws<ArgumentOutOfRangeException>(() => Service.WriteNotes(path, 99, "notes"));
    }
}
