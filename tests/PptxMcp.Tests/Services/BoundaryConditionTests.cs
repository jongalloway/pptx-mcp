namespace PptxMcp.Tests.Services;

/// <summary>
/// Boundary condition tests for PresentationService.
/// Covers slide index boundaries, shape index boundaries, table cell coordinates,
/// empty presentations, slides with no shapes, long text, Unicode/special characters,
/// and duplicate shape names across slides.
/// </summary>
public class BoundaryConditionTests : IDisposable
{
    private readonly PresentationService _service = new();
    private readonly List<string> _tempFiles = [];

    public void Dispose()
    {
        foreach (var file in _tempFiles)
            if (File.Exists(file)) File.Delete(file);
    }

    private string CreatePresentation(params TestSlideDefinition[] slides)
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        _tempFiles.Add(path);
        TestPptxHelper.CreatePresentation(path, slides);
        return path;
    }

    private string CreateMinimalPresentation(string? titleText = "Test Slide")
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        _tempFiles.Add(path);
        TestPptxHelper.CreateMinimalPresentation(path, titleText);
        return path;
    }

    // ────────────────────────────────────────────────────────────────────────
    // Slide index boundaries
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void GetSlideContent_FirstSlide_ReturnsContent()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "First" },
            new TestSlideDefinition { TitleText = "Second" });
        var content = _service.GetSlideContent(path, 0);
        Assert.Equal(0, content.SlideIndex);
    }

    [Fact]
    public void GetSlideContent_LastSlide_ReturnsContent()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "First" },
            new TestSlideDefinition { TitleText = "Second" },
            new TestSlideDefinition { TitleText = "Third" });
        var content = _service.GetSlideContent(path, 2);
        Assert.Equal(2, content.SlideIndex);
    }

    [Fact]
    public void GetSlideContent_OnePastLast_ThrowsOutOfRange()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "Only" });
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.GetSlideContent(path, 1));
    }

    [Fact]
    public void GetSlideContent_NegativeIndex_ThrowsOutOfRange()
    {
        var path = CreateMinimalPresentation();
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.GetSlideContent(path, -1));
    }

    [Fact]
    public void GetSlideXml_FirstSlide_ReturnsXml()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });
        var xml = _service.GetSlideXml(path, 0);
        Assert.Contains("sld", xml);
    }

    [Fact]
    public void GetSlideXml_LastSlide_ReturnsXml()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });
        var xml = _service.GetSlideXml(path, 1);
        Assert.Contains("sld", xml);
    }

    [Fact]
    public void GetSlideXml_NegativeIndex_ThrowsOutOfRange()
    {
        var path = CreateMinimalPresentation();
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.GetSlideXml(path, -1));
    }

    [Fact]
    public void WriteNotes_FirstSlide_Succeeds()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });
        _service.WriteNotes(path, 0, "Notes on first");
        var slides = _service.GetSlides(path);
        Assert.Equal("Notes on first", slides[0].Notes);
    }

    [Fact]
    public void WriteNotes_LastSlide_Succeeds()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" },
            new TestSlideDefinition { TitleText = "C" });
        _service.WriteNotes(path, 2, "Notes on last");
        var slides = _service.GetSlides(path);
        Assert.Equal("Notes on last", slides[2].Notes);
    }

    [Fact]
    public void WriteNotes_NegativeIndex_ThrowsOutOfRange()
    {
        var path = CreateMinimalPresentation();
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.WriteNotes(path, -1, "notes"));
    }

    [Fact]
    public void MoveSlide_FirstToLast_ReordersCorrectly()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" },
            new TestSlideDefinition { TitleText = "C" });
        _service.MoveSlide(path, 1, 3);
        var titles = _service.GetSlides(path).Select(s => s.Title).ToArray();
        Assert.Equal(["B", "C", "A"], titles);
    }

    [Fact]
    public void MoveSlide_LastToFirst_ReordersCorrectly()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" },
            new TestSlideDefinition { TitleText = "C" });
        _service.MoveSlide(path, 3, 1);
        var titles = _service.GetSlides(path).Select(s => s.Title).ToArray();
        Assert.Equal(["C", "A", "B"], titles);
    }

    [Fact]
    public void MoveSlide_ZeroSlideNumber_ThrowsOutOfRange()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.MoveSlide(path, 0, 1));
    }

    [Fact]
    public void MoveSlide_OnePastLastSlideNumber_ThrowsOutOfRange()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.MoveSlide(path, 3, 1));
    }

    [Fact]
    public void DeleteSlide_FirstSlide_RemovesCorrectSlide()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });
        _service.DeleteSlide(path, 1);
        var titles = _service.GetSlides(path).Select(s => s.Title).ToArray();
        Assert.Equal(["B"], titles);
    }

    [Fact]
    public void DeleteSlide_LastSlide_RemovesCorrectSlide()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });
        _service.DeleteSlide(path, 2);
        var titles = _service.GetSlides(path).Select(s => s.Title).ToArray();
        Assert.Equal(["A"], titles);
    }

    [Fact]
    public void DeleteSlide_ZeroSlideNumber_ThrowsOutOfRange()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.DeleteSlide(path, 0));
    }

    [Fact]
    public void DeleteSlide_OnePastLast_ThrowsOutOfRange()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });
        Assert.Throws<ArgumentOutOfRangeException>(() => _service.DeleteSlide(path, 3));
    }

    // ────────────────────────────────────────────────────────────────────────
    // Shape / placeholder index boundaries
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void UpdateTextPlaceholder_ZeroIndex_UpdatesFirstPlaceholder()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Original",
            TextShapes = [new TestTextShapeDefinition { Paragraphs = ["Body text"], PlaceholderType = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body }]
        });
        _service.UpdateTextPlaceholder(path, 0, 0, "Updated");
        var content = _service.GetSlideContent(path, 0);
        var placeholder = content.Shapes.First(s => s.IsPlaceholder);
        Assert.Equal("Updated", placeholder.Text);
    }

    [Fact]
    public void UpdateTextPlaceholder_NegativeIndex_ThrowsOutOfRange()
    {
        var path = CreateMinimalPresentation();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _service.UpdateTextPlaceholder(path, 0, -1, "text"));
    }

    [Fact]
    public void UpdateTextPlaceholder_OnePastMax_ThrowsOutOfRange()
    {
        var path = CreateMinimalPresentation();
        // Single slide has 1 placeholder (title); index 1 is one past max
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _service.UpdateTextPlaceholder(path, 0, 1, "text"));
    }

    [Fact]
    public void UpdateSlideData_SlideNumber0_ReturnsFailure()
    {
        var path = CreateMinimalPresentation();
        var result = _service.UpdateSlideData(path, 0, "Title 1", null, "text");
        Assert.False(result.Success);
        Assert.Contains("1 or greater", result.Message);
    }

    [Fact]
    public void UpdateSlideData_SlideNumberOnePastLast_ReturnsFailure()
    {
        var path = CreateMinimalPresentation();
        var result = _service.UpdateSlideData(path, 2, "Title 1", null, "text");
        Assert.False(result.Success);
        Assert.Contains("out of range", result.Message);
    }

    [Fact]
    public void UpdateSlideData_NegativePlaceholderIndex_ReturnsFailure()
    {
        var path = CreateMinimalPresentation();
        var result = _service.UpdateSlideData(path, 1, null, -1, "text");
        Assert.False(result.Success);
        Assert.Contains("zero or greater", result.Message);
    }

    [Fact]
    public void UpdateSlideData_NeitherShapeNameNorIndex_ReturnsFailure()
    {
        var path = CreateMinimalPresentation();
        var result = _service.UpdateSlideData(path, 1, null, null, "text");
        Assert.False(result.Success);
        Assert.Contains("shapeName or placeholderIndex", result.Message);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Table cell coordinate boundaries
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void UpdateTable_CellAtOrigin_Updates()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Table Slide",
            Tables = [new TestTableDefinition
            {
                Name = "Grid",
                Rows = [["A", "B"], ["C", "D"]]
            }]
        });

        var result = _service.UpdateTable(path, 1,
            [new TableCellUpdate(0, 0, "Updated")], tableName: "Grid");

        Assert.True(result.Success);
        Assert.Equal(1, result.CellsUpdated);

        var content = _service.GetSlideContent(path, 0);
        var table = content.Shapes.First(s => s.ShapeType == "Table");
        Assert.Equal("Updated", table.TableRows![0][0]);
    }

    [Fact]
    public void UpdateTable_CellAtMaxRowCol_Updates()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Table Slide",
            Tables = [new TestTableDefinition
            {
                Name = "Grid",
                Rows = [["A", "B", "C"], ["D", "E", "F"], ["G", "H", "I"]]
            }]
        });

        var result = _service.UpdateTable(path, 1,
            [new TableCellUpdate(2, 2, "Updated")], tableName: "Grid");

        Assert.True(result.Success);
        Assert.Equal(1, result.CellsUpdated);

        var content = _service.GetSlideContent(path, 0);
        var table = content.Shapes.First(s => s.ShapeType == "Table");
        Assert.Equal("Updated", table.TableRows![2][2]);
    }

    [Fact]
    public void UpdateTable_NegativeRow_SkipsCell()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Table Slide",
            Tables = [new TestTableDefinition
            {
                Name = "Grid",
                Rows = [["A"]]
            }]
        });

        var result = _service.UpdateTable(path, 1,
            [new TableCellUpdate(-1, 0, "X")], tableName: "Grid");

        Assert.True(result.Success);
        Assert.Equal(0, result.CellsUpdated);
        Assert.Equal(1, result.CellsSkipped);
    }

    [Fact]
    public void UpdateTable_NegativeColumn_SkipsCell()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Table Slide",
            Tables = [new TestTableDefinition
            {
                Name = "Grid",
                Rows = [["A"]]
            }]
        });

        var result = _service.UpdateTable(path, 1,
            [new TableCellUpdate(0, -1, "X")], tableName: "Grid");

        Assert.True(result.Success);
        Assert.Equal(0, result.CellsUpdated);
        Assert.Equal(1, result.CellsSkipped);
    }

    [Fact]
    public void UpdateTable_RowOnePastMax_SkipsCell()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Table Slide",
            Tables = [new TestTableDefinition
            {
                Name = "Grid",
                Rows = [["A", "B"], ["C", "D"]]
            }]
        });

        var result = _service.UpdateTable(path, 1,
            [new TableCellUpdate(2, 0, "X")], tableName: "Grid");

        Assert.True(result.Success);
        Assert.Equal(0, result.CellsUpdated);
        Assert.Equal(1, result.CellsSkipped);
    }

    [Fact]
    public void UpdateTable_ColumnOnePastMax_SkipsCell()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Table Slide",
            Tables = [new TestTableDefinition
            {
                Name = "Grid",
                Rows = [["A", "B"], ["C", "D"]]
            }]
        });

        var result = _service.UpdateTable(path, 1,
            [new TableCellUpdate(0, 2, "X")], tableName: "Grid");

        Assert.True(result.Success);
        Assert.Equal(0, result.CellsUpdated);
        Assert.Equal(1, result.CellsSkipped);
    }

    [Fact]
    public void UpdateTable_MixOfValidAndInvalidCoords_ReportsCorrectCounts()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Table Slide",
            Tables = [new TestTableDefinition
            {
                Name = "Grid",
                Rows = [["A", "B"], ["C", "D"]]
            }]
        });

        var result = _service.UpdateTable(path, 1,
            [
                new TableCellUpdate(0, 0, "Valid"),
                new TableCellUpdate(-1, 0, "BadRow"),
                new TableCellUpdate(0, -1, "BadCol"),
                new TableCellUpdate(1, 1, "AlsoValid"),
                new TableCellUpdate(99, 99, "WayOff")
            ], tableName: "Grid");

        Assert.True(result.Success);
        Assert.Equal(2, result.CellsUpdated);
        Assert.Equal(3, result.CellsSkipped);
    }

    [Fact]
    public void UpdateTable_TableIndex_NegativeIndex_ThrowsOutOfRange()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Table Slide",
            Tables = [new TestTableDefinition { Rows = [["A"]] }]
        });

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _service.UpdateTable(path, 1, [new TableCellUpdate(0, 0, "X")], tableIndex: -1));
    }

    [Fact]
    public void UpdateTable_TableIndex_OnePastMax_ThrowsOutOfRange()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Table Slide",
            Tables = [new TestTableDefinition { Rows = [["A"]] }]
        });

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _service.UpdateTable(path, 1, [new TableCellUpdate(0, 0, "X")], tableIndex: 1));
    }

    // ────────────────────────────────────────────────────────────────────────
    // Empty presentations / slides with no shapes
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void GetSlides_EmptyPresentation_ReturnsEmptyList()
    {
        var path = CreatePresentation(); // no slides at all
        var slides = _service.GetSlides(path);
        Assert.Empty(slides);
    }

    [Fact]
    public void GetAllSlideContents_EmptyPresentation_ReturnsEmptyList()
    {
        var path = CreatePresentation();
        var contents = _service.GetAllSlideContents(path);
        Assert.Empty(contents);
    }

    [Fact]
    public void ExtractTalkingPoints_EmptyPresentation_ReturnsEmptyList()
    {
        var path = CreatePresentation();
        var points = _service.ExtractTalkingPoints(path);
        Assert.Empty(points);
    }

    [Fact]
    public void UpdateSlideData_EmptyPresentation_ReturnsFailure()
    {
        var path = CreatePresentation();
        var result = _service.UpdateSlideData(path, 1, "shape", null, "text");
        Assert.False(result.Success);
        Assert.Contains("no slides", result.Message);
    }

    [Fact]
    public void GetSlideContent_SlideWithNoShapes_ReturnsEmptyShapes()
    {
        var path = CreatePresentation(new TestSlideDefinition()); // no title, no text shapes
        var content = _service.GetSlideContent(path, 0);
        Assert.Empty(content.Shapes);
    }

    [Fact]
    public void ExtractTalkingPoints_SlideWithNoShapes_ReturnsEmptyPoints()
    {
        var path = CreatePresentation(new TestSlideDefinition());
        var points = _service.ExtractTalkingPoints(path);
        Assert.Single(points);
        Assert.Empty(points[0].Points);
    }

    [Fact]
    public void UpdateSlideData_SlideWithNoTextShapes_ReturnsFailure()
    {
        // Image-only slide has no text shapes
        var path = CreatePresentation(new TestSlideDefinition { IncludeImage = true });
        var result = _service.UpdateSlideData(path, 1, "SomeShape", null, "text");
        Assert.False(result.Success);
        Assert.Contains("does not contain any text-capable shapes", result.Message);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Very long text values
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void UpdateSlideData_VeryLongText_PreservesFullContent()
    {
        var longText = new string('X', 50_000);
        var path = CreateMinimalPresentation("Short");
        var result = _service.UpdateSlideData(path, 1, "Title 1", null, longText);

        Assert.True(result.Success);
        var content = _service.GetSlideContent(path, 0);
        var shape = content.Shapes.First(s => s.Name == "Title 1");
        Assert.Equal(50_000, shape.Text!.Length);
        Assert.True(shape.Text!.All(c => c == 'X'));
    }

    [Fact]
    public void InsertTable_VeryLongCellText_PreservesFullContent()
    {
        var longValue = new string('Y', 10_000);
        var path = CreateMinimalPresentation();
        var result = _service.InsertTable(path, 1,
            ["Header"],
            [new[] { longValue }]);

        Assert.True(result.Success);
        var content = _service.GetSlideContent(path, 0);
        var table = content.Shapes.First(s => s.ShapeType == "Table");
        Assert.Equal(longValue, table.TableRows![1][0]);
    }

    [Fact]
    public void WriteNotes_VeryLongText_PreservesFullContent()
    {
        var longText = new string('Z', 50_000);
        var path = CreateMinimalPresentation();
        _service.WriteNotes(path, 0, longText);
        var slides = _service.GetSlides(path);
        Assert.Equal(50_000, slides[0].Notes!.Length);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Unicode and special characters
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void UpdateSlideData_UnicodeShapeName_MatchesCorrectly()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Dashboard",
            TextShapes = [new TestTextShapeDefinition
            {
                Name = "売上高ラベル",
                Paragraphs = ["Revenue"]
            }]
        });

        var result = _service.UpdateSlideData(path, 1, "売上高ラベル", null, "収益");
        Assert.True(result.Success);

        var content = _service.GetSlideContent(path, 0);
        var shape = content.Shapes.First(s => s.Name == "売上高ラベル");
        Assert.Equal("収益", shape.Text);
    }

    [Fact]
    public void UpdateSlideData_EmojiText_PreservesEmoji()
    {
        var path = CreateMinimalPresentation("Placeholder");
        var result = _service.UpdateSlideData(path, 1, "Title 1", null, "🚀 Launch 🎉 Party 🌍");
        Assert.True(result.Success);

        var content = _service.GetSlideContent(path, 0);
        var shape = content.Shapes.First(s => s.Name == "Title 1");
        Assert.Equal("🚀 Launch 🎉 Party 🌍", shape.Text);
    }

    [Fact]
    public void InsertTable_UnicodeAndSpecialCharsInCells_Preserves()
    {
        var path = CreateMinimalPresentation();
        var headers = new[] { "名前", "Wert" };
        var rows = new[] { new[] { "Ñoño <>&\"'", "Ü — €£¥" } };

        var result = _service.InsertTable(path, 1, headers, rows);
        Assert.True(result.Success);

        var content = _service.GetSlideContent(path, 0);
        var table = content.Shapes.First(s => s.ShapeType == "Table");
        Assert.Equal("名前", table.TableRows![0][0]);
        Assert.Equal("Wert", table.TableRows[0][1]);
        Assert.Equal("Ñoño <>&\"'", table.TableRows[1][0]);
        Assert.Equal("Ü — €£¥", table.TableRows[1][1]);
    }

    [Fact]
    public void UpdateTable_UnicodeValue_PreservesContent()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Table",
            Tables = [new TestTableDefinition
            {
                Name = "Data",
                Rows = [["Placeholder"]]
            }]
        });

        var result = _service.UpdateTable(path, 1,
            [new TableCellUpdate(0, 0, "中文数据 — Données")], tableName: "Data");

        Assert.True(result.Success);
        var content = _service.GetSlideContent(path, 0);
        var table = content.Shapes.First(s => s.ShapeType == "Table");
        Assert.Equal("中文数据 — Données", table.TableRows![0][0]);
    }

    [Fact]
    public void WriteNotes_UnicodeText_PreservesContent()
    {
        var path = CreateMinimalPresentation();
        _service.WriteNotes(path, 0, "Заметки докладчика — प्रस्तुतकर्ता");
        var slides = _service.GetSlides(path);
        Assert.Equal("Заметки докладчика — प्रस्तुतकर्ता", slides[0].Notes);
    }

    [Fact]
    public void ExportMarkdown_UnicodeContent_PreservesInOutput()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "日本語タイトル",
            TextShapes = [new TestTextShapeDefinition
            {
                Name = "Body",
                Paragraphs = ["Ελληνικά κείμενο"],
                PlaceholderType = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body
            }]
        });

        var result = _service.ExportMarkdown(path);
        Assert.Contains("日本語タイトル", result.Markdown);
        Assert.Contains("Ελληνικά κείμενο", result.Markdown);
    }

    [Fact]
    public void UpdateSlideData_SpecialXmlChars_PreservesContent()
    {
        var path = CreateMinimalPresentation("Placeholder");
        var specialText = "Revenue > $1M & <strong>bold</strong> \"quoted\"";
        var result = _service.UpdateSlideData(path, 1, "Title 1", null, specialText);
        Assert.True(result.Success);

        var content = _service.GetSlideContent(path, 0);
        var shape = content.Shapes.First(s => s.Name == "Title 1");
        Assert.Equal(specialText, shape.Text);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Multiple slides with same shape names
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void UpdateSlideData_DuplicateShapeNamesAcrossSlides_UpdatesCorrectSlide()
    {
        var path = CreatePresentation(
            new TestSlideDefinition
            {
                TitleText = "Slide 1",
                TextShapes = [new TestTextShapeDefinition
                {
                    Name = "Value",
                    Paragraphs = ["100"]
                }]
            },
            new TestSlideDefinition
            {
                TitleText = "Slide 2",
                TextShapes = [new TestTextShapeDefinition
                {
                    Name = "Value",
                    Paragraphs = ["200"]
                }]
            });

        // Update shape "Value" on slide 2 only
        var result = _service.UpdateSlideData(path, 2, "Value", null, "999");
        Assert.True(result.Success);

        // Slide 1 should be unchanged
        var slide1 = _service.GetSlideContent(path, 0);
        var shape1 = slide1.Shapes.First(s => s.Name == "Value");
        Assert.Equal("100", shape1.Text);

        // Slide 2 should be updated
        var slide2 = _service.GetSlideContent(path, 1);
        var shape2 = slide2.Shapes.First(s => s.Name == "Value");
        Assert.Equal("999", shape2.Text);
    }

    [Fact]
    public void BatchUpdate_DuplicateShapeNamesAcrossSlides_UpdatesEachSlideIndependently()
    {
        var path = CreatePresentation(
            new TestSlideDefinition
            {
                TitleText = "Slide 1",
                TextShapes = [new TestTextShapeDefinition
                {
                    Name = "Metric",
                    Paragraphs = ["Old1"]
                }]
            },
            new TestSlideDefinition
            {
                TitleText = "Slide 2",
                TextShapes = [new TestTextShapeDefinition
                {
                    Name = "Metric",
                    Paragraphs = ["Old2"]
                }]
            });

        var result = _service.BatchUpdate(path, [
            new BatchUpdateMutation(1, "Metric", "New1"),
            new BatchUpdateMutation(2, "Metric", "New2")
        ]);

        Assert.Equal(2, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);

        var slide1 = _service.GetSlideContent(path, 0);
        Assert.Equal("New1", slide1.Shapes.First(s => s.Name == "Metric").Text);

        var slide2 = _service.GetSlideContent(path, 1);
        Assert.Equal("New2", slide2.Shapes.First(s => s.Name == "Metric").Text);
    }

    // ────────────────────────────────────────────────────────────────────────
    // ReorderSlides boundary conditions
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void ReorderSlides_IdentityPermutation_PreservesOrder()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" },
            new TestSlideDefinition { TitleText = "C" });
        _service.ReorderSlides(path, [1, 2, 3]);
        var titles = _service.GetSlides(path).Select(s => s.Title).ToArray();
        Assert.Equal(["A", "B", "C"], titles);
    }

    [Fact]
    public void ReorderSlides_ReverseOrder_Works()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" },
            new TestSlideDefinition { TitleText = "C" });
        _service.ReorderSlides(path, [3, 2, 1]);
        var titles = _service.GetSlides(path).Select(s => s.Title).ToArray();
        Assert.Equal(["C", "B", "A"], titles);
    }

    [Fact]
    public void ReorderSlides_WrongElementCount_ThrowsArgumentException()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });
        Assert.Throws<ArgumentException>(() => _service.ReorderSlides(path, [1]));
    }

    [Fact]
    public void ReorderSlides_InvalidValues_ThrowsArgumentException()
    {
        var path = CreatePresentation(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });
        Assert.Throws<ArgumentException>(() => _service.ReorderSlides(path, [0, 1]));
    }

    // ────────────────────────────────────────────────────────────────────────
    // BatchUpdate boundary conditions
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void BatchUpdate_EmptyMutationsList_ReturnsZeroCounts()
    {
        var path = CreateMinimalPresentation();
        var result = _service.BatchUpdate(path, []);
        Assert.Equal(0, result.TotalMutations);
        Assert.Equal(0, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
    }

    [Fact]
    public void BatchUpdate_MixOfValidAndInvalidSlideNumbers_ReportsCorrectly()
    {
        var path = CreateMinimalPresentation("Title");
        var result = _service.BatchUpdate(path, [
            new BatchUpdateMutation(1, "Title 1", "Updated"),
            new BatchUpdateMutation(99, "Title 1", "Nope")
        ]);

        Assert.Equal(2, result.TotalMutations);
        Assert.Equal(1, result.SuccessCount);
        Assert.Equal(1, result.FailureCount);
        Assert.True(result.Results[0].Success);
        Assert.False(result.Results[1].Success);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Single-slide presentation edge cases
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void DeleteSlide_OnlySlide_ThrowsInvalidOperation()
    {
        var path = CreateMinimalPresentation();
        Assert.Throws<InvalidOperationException>(() => _service.DeleteSlide(path, 1));
    }

    [Fact]
    public void MoveSlide_SingleSlide_SamePosition_IsNoOp()
    {
        var path = CreateMinimalPresentation("Solo");
        _service.MoveSlide(path, 1, 1);
        var titles = _service.GetSlides(path).Select(s => s.Title).ToArray();
        Assert.Equal(["Solo"], titles);
    }

    [Fact]
    public void ReorderSlides_SingleSlide_IdentityPermutation_Works()
    {
        var path = CreateMinimalPresentation("Solo");
        _service.ReorderSlides(path, [1]);
        var titles = _service.GetSlides(path).Select(s => s.Title).ToArray();
        Assert.Equal(["Solo"], titles);
    }

    // ────────────────────────────────────────────────────────────────────────
    // InsertTable on no-slides presentation
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void InsertTable_EmptyPresentation_ThrowsOnSlideAccess()
    {
        var path = CreatePresentation();
        Assert.ThrowsAny<Exception>(() =>
            _service.InsertTable(path, 1, ["H"], [new[] { "V" }]));
    }

    [Fact]
    public void UpdateTable_NoTablesOnSlide_ThrowsInvalidOperation()
    {
        var path = CreateMinimalPresentation();
        Assert.Throws<InvalidOperationException>(() =>
            _service.UpdateTable(path, 1, [new TableCellUpdate(0, 0, "X")]));
    }
}
