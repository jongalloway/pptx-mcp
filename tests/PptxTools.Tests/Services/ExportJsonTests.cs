using DocumentFormat.OpenXml.Packaging;

namespace PptxTools.Tests.Services;

/// <summary>
/// Service-level tests for ExportJson (Issue #128 — Export presentation to structured JSON).
/// Covers Full, SlidesOnly, MetadataOnly, SchemaOnly actions, text/table/image/chart/notes content,
/// metadata extraction, minimal presentations, error handling, and structural invariants.
/// </summary>
[Trait("Category", "Unit")]
public class ExportJsonTests : PptxTestBase
{
    // ────────────────────────────────────────────────────────
    // Fixture helpers
    // ────────────────────────────────────────────────────────

    private string CreatePptxWithMetadata(
        string? title = null,
        string? creator = null,
        string? subject = null,
        string? keywords = null,
        string? description = null,
        string? category = null,
        params TestSlideDefinition[] slides)
    {
        var defs = slides.Length > 0 ? slides : [new TestSlideDefinition { TitleText = "Slide 1" }];
        var path = CreatePptxWithSlides(defs);

        using var doc = PresentationDocument.Open(path, true);
        if (title is not null) doc.PackageProperties.Title = title;
        if (creator is not null) doc.PackageProperties.Creator = creator;
        if (subject is not null) doc.PackageProperties.Subject = subject;
        if (keywords is not null) doc.PackageProperties.Keywords = keywords;
        if (description is not null) doc.PackageProperties.Description = description;
        if (category is not null) doc.PackageProperties.Category = category;

        return path;
    }

    // ────────────────────────────────────────────────────────
    // Full export: minimal presentation
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_MinimalPresentation_ReturnsSuccess()
    {
        var path = CreateMinimalPptx();

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.True(result.Success);
    }

    [Fact]
    public void ExportJson_MinimalPresentation_SlideCountIsOne()
    {
        var path = CreateMinimalPptx();

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.Equal(1, result.SlideCount);
    }

    [Fact]
    public void ExportJson_MinimalPresentation_ActionIsFull()
    {
        var path = CreateMinimalPptx();

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.Equal("Full", result.Action);
    }

    [Fact]
    public void ExportJson_MinimalPresentation_FilePathMatches()
    {
        var path = CreateMinimalPptx();

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.Equal(path, result.FilePath);
    }

    [Fact]
    public void ExportJson_MinimalPresentation_SlidesNotNull()
    {
        var path = CreateMinimalPptx();

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        Assert.Single(result.Slides);
    }

    [Fact]
    public void ExportJson_MinimalPresentation_MetadataNotNull()
    {
        var path = CreateMinimalPptx();

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Metadata);
    }

    [Fact]
    public void ExportJson_MinimalPresentation_SchemaIsNull()
    {
        var path = CreateMinimalPptx();

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.Null(result.Schema);
    }

    // ────────────────────────────────────────────────────────
    // Full export: text shapes
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_WithTextShapes_ShapesContainText()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "My Title",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Content 1",
                        Paragraphs = ["Hello World", "Second paragraph"]
                    }
                ]
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.True(result.Success);
        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        Assert.True(slide.Shapes.Count >= 2, "Expected at least title + content shape");
    }

    [Fact]
    public void ExportJson_WithTitle_SlideHasTitle()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Quarterly Review" });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        Assert.Equal("Quarterly Review", slide.Title);
    }

    [Fact]
    public void ExportJson_WithMultipleTextShapes_AllShapesExported()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Slide Title",
                TextShapes =
                [
                    new TestTextShapeDefinition { Name = "Shape A", Paragraphs = ["Alpha"] },
                    new TestTextShapeDefinition { Name = "Shape B", Paragraphs = ["Beta"] },
                    new TestTextShapeDefinition { Name = "Shape C", Paragraphs = ["Gamma"] }
                ]
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        // Title + 3 text shapes = at least 4
        Assert.True(slide.Shapes.Count >= 4, $"Expected ≥4 shapes, got {slide.Shapes.Count}");
    }

    [Fact]
    public void ExportJson_TextShape_HasParagraphs()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Bullets",
                        Paragraphs = ["Point one", "Point two", "Point three"]
                    }
                ]
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        var shape = slide.Shapes.FirstOrDefault(s => s.Name == "Bullets");
        Assert.NotNull(shape);
        Assert.NotNull(shape.Paragraphs);
        Assert.Equal(3, shape.Paragraphs.Count);
    }

    // ────────────────────────────────────────────────────────
    // Full export: tables (embedded in ShapeExport.Table)
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_WithTable_ShapeHasTableData()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Data Slide",
                Tables =
                [
                    new TestTableDefinition
                    {
                        Name = "Metrics Table",
                        Rows = [["Q1", "100"], ["Q2", "200"], ["Q3", "300"]]
                    }
                ]
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.True(result.Success);
        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        var tableShape = slide.Shapes.FirstOrDefault(s => s.Table is not null);
        Assert.NotNull(tableShape);
        Assert.Equal("Table", tableShape.ShapeType);
    }

    [Fact]
    public void ExportJson_WithTable_TableHasCorrectDimensions()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                Tables =
                [
                    new TestTableDefinition
                    {
                        Name = "Small Table",
                        Rows = [["A", "B"], ["C", "D"]]
                    }
                ]
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        var tableShape = slide.Shapes.FirstOrDefault(s => s.Table is not null);
        Assert.NotNull(tableShape);
        Assert.NotNull(tableShape.Table);
        Assert.Equal(2, tableShape.Table.RowCount);
        Assert.Equal(2, tableShape.Table.ColumnCount);
    }

    [Fact]
    public void ExportJson_WithTable_CellsPreserved()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                Tables =
                [
                    new TestTableDefinition
                    {
                        Name = "Values Table",
                        Rows = [["Revenue", "$1M"], ["Cost", "$500K"]]
                    }
                ]
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        var tableShape = slide.Shapes.FirstOrDefault(s => s.Table is not null);
        Assert.NotNull(tableShape);
        Assert.NotNull(tableShape.Table);
        Assert.NotNull(tableShape.Table.Cells);
        Assert.True(tableShape.Table.Cells.Count >= 2, "Expected at least 2 rows of cell data");
    }

    [Fact]
    public void ExportJson_WithTable_ShapeNameMatches()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                Tables =
                [
                    new TestTableDefinition
                    {
                        Name = "KPI Grid",
                        Rows = [["Metric", "Value"]]
                    }
                ]
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        var tableShape = slide.Shapes.FirstOrDefault(s => s.Table is not null);
        Assert.NotNull(tableShape);
        Assert.Equal("KPI Grid", tableShape.Name);
    }

    // ────────────────────────────────────────────────────────
    // Full export: images (embedded in ShapeExport.Image)
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_WithImage_ImagesPopulated()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Image Slide",
                IncludeImage = true
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.True(result.Success);
        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        Assert.True(slide.Images.Count > 0, "Expected at least one image via computed property");
    }

    [Fact]
    public void ExportJson_WithImage_ImageHasContentType()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Picture Slide",
                IncludeImage = true
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        var imageShape = slide.Shapes.FirstOrDefault(s => s.Image is not null);
        Assert.NotNull(imageShape);
        Assert.NotNull(imageShape.Image);
        Assert.False(string.IsNullOrWhiteSpace(imageShape.Image.ContentType));
    }

    [Fact]
    public void ExportJson_WithImage_ImageHasRelationshipId()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Pic Slide",
                IncludeImage = true
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        var imageShape = slide.Shapes.FirstOrDefault(s => s.Image is not null);
        Assert.NotNull(imageShape);
        Assert.NotNull(imageShape.Image);
        Assert.False(string.IsNullOrWhiteSpace(imageShape.Image.RelationshipId));
    }

    // ────────────────────────────────────────────────────────
    // Full export: speaker notes
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_WithSpeakerNotes_NotesExported()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Notes Slide",
                SpeakerNotesText = "Remember to pause here."
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        Assert.Equal("Remember to pause here.", slide.SpeakerNotes);
    }

    [Fact]
    public void ExportJson_WithoutSpeakerNotes_NotesIsNull()
    {
        var path = CreateMinimalPptx();

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        Assert.Null(slide.SpeakerNotes);
    }

    // ────────────────────────────────────────────────────────
    // MetadataOnly action
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_MetadataOnly_ReturnsSuccess()
    {
        var path = CreatePptxWithMetadata(title: "Annual Report", creator: "Jane Doe");

        var result = Service.ExportJson(path, ExportJsonAction.MetadataOnly);

        Assert.True(result.Success);
        Assert.Equal("MetadataOnly", result.Action);
    }

    [Fact]
    public void ExportJson_MetadataOnly_MetadataPopulated()
    {
        var path = CreatePptxWithMetadata(
            title: "Quarterly Review",
            creator: "John Smith",
            subject: "Q4 Results",
            keywords: "finance, review");

        var result = Service.ExportJson(path, ExportJsonAction.MetadataOnly);

        Assert.NotNull(result.Metadata);
        Assert.Equal("Quarterly Review", result.Metadata.Title);
        Assert.Equal("John Smith", result.Metadata.Creator);
        Assert.Equal("Q4 Results", result.Metadata.Subject);
        Assert.Equal("finance, review", result.Metadata.Keywords);
    }

    [Fact]
    public void ExportJson_MetadataOnly_SlidesNull()
    {
        var path = CreatePptxWithMetadata(title: "Test");

        var result = Service.ExportJson(path, ExportJsonAction.MetadataOnly);

        Assert.Null(result.Slides);
    }

    [Fact]
    public void ExportJson_MetadataOnly_AllMetadataFieldsAccessible()
    {
        var path = CreatePptxWithMetadata(
            title: "T",
            creator: "C",
            subject: "S",
            keywords: "K",
            description: "D",
            category: "Cat");

        var result = Service.ExportJson(path, ExportJsonAction.MetadataOnly);

        Assert.NotNull(result.Metadata);
        Assert.Equal("T", result.Metadata.Title);
        Assert.Equal("C", result.Metadata.Creator);
        Assert.Equal("S", result.Metadata.Subject);
        Assert.Equal("K", result.Metadata.Keywords);
        Assert.Equal("D", result.Metadata.Description);
        Assert.Equal("Cat", result.Metadata.Category);
    }

    // ────────────────────────────────────────────────────────
    // SlidesOnly action
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_SlidesOnly_ReturnsSuccess()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Slide A" },
            new TestSlideDefinition { TitleText = "Slide B" });

        var result = Service.ExportJson(path, ExportJsonAction.SlidesOnly);

        Assert.True(result.Success);
        Assert.Equal("SlidesOnly", result.Action);
    }

    [Fact]
    public void ExportJson_SlidesOnly_SlidesPopulated()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "First" },
            new TestSlideDefinition { TitleText = "Second" },
            new TestSlideDefinition { TitleText = "Third" });

        var result = Service.ExportJson(path, ExportJsonAction.SlidesOnly);

        Assert.NotNull(result.Slides);
        Assert.Equal(3, result.Slides.Count);
        Assert.Equal(3, result.SlideCount);
    }

    [Fact]
    public void ExportJson_SlidesOnly_MetadataNull()
    {
        var path = CreatePptxWithMetadata(title: "Ignored Title");

        var result = Service.ExportJson(path, ExportJsonAction.SlidesOnly);

        Assert.Null(result.Metadata);
    }

    [Fact]
    public void ExportJson_SlidesOnly_SlideNumbersAreOneBased()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Alpha" },
            new TestSlideDefinition { TitleText = "Beta" });

        var result = Service.ExportJson(path, ExportJsonAction.SlidesOnly);

        Assert.NotNull(result.Slides);
        Assert.Equal(1, result.Slides[0].SlideNumber);
        Assert.Equal(2, result.Slides[1].SlideNumber);
    }

    [Fact]
    public void ExportJson_SlidesOnly_SlideIndicesAreZeroBased()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });

        var result = Service.ExportJson(path, ExportJsonAction.SlidesOnly);

        Assert.NotNull(result.Slides);
        Assert.Equal(0, result.Slides[0].SlideIndex);
        Assert.Equal(1, result.Slides[1].SlideIndex);
    }

    // ────────────────────────────────────────────────────────
    // SchemaOnly action
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_SchemaOnly_ReturnsSuccess()
    {
        var result = Service.ExportJson("", ExportJsonAction.SchemaOnly);

        Assert.True(result.Success);
        Assert.Equal("SchemaOnly", result.Action);
    }

    [Fact]
    public void ExportJson_SchemaOnly_SchemaNotNull()
    {
        var result = Service.ExportJson("", ExportJsonAction.SchemaOnly);

        Assert.NotNull(result.Schema);
        Assert.Contains("PresentationExport", result.Schema);
    }

    [Fact]
    public void ExportJson_SchemaOnly_SlidesAndMetadataNull()
    {
        var result = Service.ExportJson("", ExportJsonAction.SchemaOnly);

        Assert.Null(result.Slides);
        Assert.Null(result.Metadata);
    }

    [Fact]
    public void ExportJson_SchemaOnly_SlideCountIsZero()
    {
        var result = Service.ExportJson("", ExportJsonAction.SchemaOnly);

        Assert.Equal(0, result.SlideCount);
    }

    [Fact]
    public void ExportJson_SchemaOnly_FilePathIsNull()
    {
        var result = Service.ExportJson("", ExportJsonAction.SchemaOnly);

        Assert.Null(result.FilePath);
    }

    // ────────────────────────────────────────────────────────
    // Multi-slide export
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_MultipleSlides_AllSlidesExported()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Intro" },
            new TestSlideDefinition { TitleText = "Body" },
            new TestSlideDefinition { TitleText = "Summary" },
            new TestSlideDefinition { TitleText = "Appendix" });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.Equal(4, result.SlideCount);
        Assert.NotNull(result.Slides);
        Assert.Equal(4, result.Slides.Count);
    }

    [Fact]
    public void ExportJson_MultipleSlides_SlideOrderPreserved()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "First" },
            new TestSlideDefinition { TitleText = "Second" },
            new TestSlideDefinition { TitleText = "Third" });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        Assert.Equal("First", result.Slides[0].Title);
        Assert.Equal("Second", result.Slides[1].Title);
        Assert.Equal("Third", result.Slides[2].Title);
    }

    // ────────────────────────────────────────────────────────
    // Mixed content: text + tables + images on one slide
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_MixedContent_AllContentTypesPresent()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Mixed Slide",
                TextShapes =
                [
                    new TestTextShapeDefinition { Name = "Body Text", Paragraphs = ["Content here"] }
                ],
                Tables =
                [
                    new TestTableDefinition
                    {
                        Name = "Data Table",
                        Rows = [["X", "Y"]]
                    }
                ],
                IncludeImage = true
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.True(result.Success);
        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        Assert.True(slide.Shapes.Count >= 3, $"Expected ≥3 shapes, got {slide.Shapes.Count}");
        // Tables accessible via shapes with .Table embedded
        Assert.True(slide.Shapes.Any(s => s.Table is not null), "Expected a table shape");
        // Images accessible via computed property
        Assert.True(slide.Images.Count > 0, "Expected at least one image");
    }

    // ────────────────────────────────────────────────────────
    // Error handling: file not found
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_FileNotFound_Throws()
    {
        var fakePath = @"C:\does-not-exist\export.pptx";

        Assert.ThrowsAny<Exception>(() => Service.ExportJson(fakePath, ExportJsonAction.Full));
    }

    // ────────────────────────────────────────────────────────
    // Action string matches enum
    // ────────────────────────────────────────────────────────

    [Theory]
    [InlineData(ExportJsonAction.Full, "Full")]
    [InlineData(ExportJsonAction.SlidesOnly, "SlidesOnly")]
    [InlineData(ExportJsonAction.MetadataOnly, "MetadataOnly")]
    [InlineData(ExportJsonAction.SchemaOnly, "SchemaOnly")]
    public void ExportJson_ActionString_MatchesEnumName(ExportJsonAction action, string expected)
    {
        var path = action == ExportJsonAction.SchemaOnly ? "" : CreateMinimalPptx();

        var result = Service.ExportJson(path, action);

        Assert.Equal(expected, result.Action);
    }

    // ────────────────────────────────────────────────────────
    // Full export: metadata with all fields
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_Full_MetadataIncluded()
    {
        var path = CreatePptxWithMetadata(
            title: "Board Meeting",
            creator: "CEO",
            subject: "Strategy");

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Metadata);
        Assert.Equal("Board Meeting", result.Metadata.Title);
        Assert.Equal("CEO", result.Metadata.Creator);
        Assert.Equal("Strategy", result.Metadata.Subject);
    }

    // ────────────────────────────────────────────────────────
    // Idempotency
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_Idempotent_SameResultTwice()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Stable",
                TextShapes =
                [
                    new TestTextShapeDefinition { Name = "Body", Paragraphs = ["Consistent"] }
                ]
            });

        var result1 = Service.ExportJson(path, ExportJsonAction.Full);
        var result2 = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.Equal(result1.SlideCount, result2.SlideCount);
        Assert.Equal(result1.Success, result2.Success);
        Assert.Equal(result1.Action, result2.Action);
    }

    // ────────────────────────────────────────────────────────
    // Speaker notes across multiple slides
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_MultipleSlides_SpeakerNotesPerSlide()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Slide 1",
                SpeakerNotesText = "Notes for slide 1"
            },
            new TestSlideDefinition
            {
                TitleText = "Slide 2",
                SpeakerNotesText = "Notes for slide 2"
            },
            new TestSlideDefinition
            {
                TitleText = "Slide 3"
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        Assert.Equal(3, result.Slides.Count);
        Assert.Equal("Notes for slide 1", result.Slides[0].SpeakerNotes);
        Assert.Equal("Notes for slide 2", result.Slides[1].SpeakerNotes);
        Assert.Null(result.Slides[2].SpeakerNotes);
    }

    // ────────────────────────────────────────────────────────
    // Shape name preservation
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_WithNamedShapes_NamesPreserved()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes =
                [
                    new TestTextShapeDefinition { Name = "Revenue Label", Paragraphs = ["$1M"] },
                    new TestTextShapeDefinition { Name = "Cost Label", Paragraphs = ["$500K"] }
                ]
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        Assert.Contains(slide.Shapes, s => s.Name == "Revenue Label");
        Assert.Contains(slide.Shapes, s => s.Name == "Cost Label");
    }

    // ────────────────────────────────────────────────────────
    // SlideCount invariant
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_Full_SlideCountMatchesSlidesLength()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" },
            new TestSlideDefinition { TitleText = "C" });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        Assert.Equal(result.SlideCount, result.Slides.Count);
    }

    // ────────────────────────────────────────────────────────
    // Slide dimensions
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_Full_SlideHasDimensions()
    {
        var path = CreateMinimalPptx();

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        Assert.True(slide.SlideWidthEmu > 0, "Expected positive slide width");
        Assert.True(slide.SlideHeightEmu > 0, "Expected positive slide height");
    }

    // ────────────────────────────────────────────────────────
    // Charts export (via computed property)
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_WithChart_ChartsPopulated()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "Chart Slide",
                Charts =
                [
                    new TestChartDefinition
                    {
                        Name = "Revenue Chart",
                        ChartType = "Column",
                        Categories = ["Q1", "Q2", "Q3"],
                        Series =
                        [
                            new TestSeriesDefinition { Name = "Revenue", Values = [100, 200, 300] }
                        ]
                    }
                ]
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        Assert.True(slide.Charts.Count > 0, "Expected at least one chart via computed property");
    }

    [Fact]
    public void ExportJson_WithChart_ChartHasSeriesData()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                Charts =
                [
                    new TestChartDefinition
                    {
                        Name = "Sales Chart",
                        ChartType = "Bar",
                        Categories = ["Jan", "Feb"],
                        Series =
                        [
                            new TestSeriesDefinition { Name = "Sales", Values = [50, 75] }
                        ]
                    }
                ]
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        var chartShape = slide.Shapes.FirstOrDefault(s => s.Chart is not null);
        Assert.NotNull(chartShape);
        Assert.NotNull(chartShape.Chart);
        Assert.Equal(1, chartShape.Chart.SeriesCount);
        Assert.Single(chartShape.Chart.Series);
    }

    // ────────────────────────────────────────────────────────
    // Shape type classification
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ExportJson_TextShape_HasTextShapeType()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TextShapes =
                [
                    new TestTextShapeDefinition { Name = "MyText", Paragraphs = ["Hello"] }
                ]
            });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        var shape = slide.Shapes.FirstOrDefault(s => s.Name == "MyText");
        Assert.NotNull(shape);
        Assert.Equal("Text", shape.ShapeType);
    }

    [Fact]
    public void ExportJson_ImageShape_HasPictureShapeType()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.ExportJson(path, ExportJsonAction.Full);

        Assert.NotNull(result.Slides);
        var slide = Assert.Single(result.Slides);
        var pictureShape = slide.Shapes.FirstOrDefault(s => s.ShapeType == "Picture");
        Assert.NotNull(pictureShape);
    }
}
