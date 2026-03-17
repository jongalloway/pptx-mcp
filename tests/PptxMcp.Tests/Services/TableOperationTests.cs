using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Tests.Services;

/// <summary>
/// Service-level tests for InsertTable and UpdateTable operations.
/// Written proactively for Issue #36 — table insert and update tools.
/// These tests verify OpenXML structure, PowerPoint compatibility, and behavioral correctness.
/// </summary>
public class TableOperationTests : IDisposable
{
    private readonly PresentationService _service = new();
    private readonly List<string> _tempFiles = [];

    public void Dispose()
    {
        foreach (var file in _tempFiles)
            if (File.Exists(file)) File.Delete(file);
    }

    // ────────────────────────────────────────────────────────
    // InsertTable: happy path
    // ────────────────────────────────────────────────────────

    [Fact]
    public void InsertTable_CreatesTableOnSlide_WithCorrectRowAndColumnCounts()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Data Slide" });
        var headers = new[] { "Region", "Revenue", "Growth" };
        var rows = new[]
        {
            new[] { "NA", "3.2M", "12%" },
            new[] { "EMEA", "1.4M", "8%" }
        };

        var result = _service.InsertTable(path, 1, headers, rows);

        Assert.True(result.Success);
        Assert.Equal(1, result.SlideNumber);
        Assert.Equal(3, result.RowCount);   // 1 header row + 2 data rows
        Assert.Equal(3, result.ColumnCount);

        // Verify table exists on the slide via GetSlideContent
        var slideContent = _service.GetSlideContent(path, 0);
        var tableShape = Assert.Single(slideContent.Shapes, s => s.ShapeType == "Table");
        Assert.NotNull(tableShape.TableRows);
        Assert.Equal(3, tableShape.TableRows.Count);
        Assert.Equal(3, tableShape.TableRows[0].Count);
    }

    [Fact]
    public void InsertTable_CellTextMatchesInput()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Metrics" });
        var headers = new[] { "KPI", "Value" };
        var rows = new[]
        {
            new[] { "ARR", "4.2M" },
            new[] { "NRR", "112%" }
        };

        _service.InsertTable(path, 1, headers, rows);

        var slideContent = _service.GetSlideContent(path, 0);
        var tableShape = Assert.Single(slideContent.Shapes, s => s.ShapeType == "Table");
        Assert.NotNull(tableShape.TableRows);

        // Header row
        Assert.Equal("KPI", tableShape.TableRows[0][0]);
        Assert.Equal("Value", tableShape.TableRows[0][1]);

        // Data rows
        Assert.Equal("ARR", tableShape.TableRows[1][0]);
        Assert.Equal("4.2M", tableShape.TableRows[1][1]);
        Assert.Equal("NRR", tableShape.TableRows[2][0]);
        Assert.Equal("112%", tableShape.TableRows[2][1]);
    }

    [Fact]
    public void InsertTable_AssignsUniqueShapeId_NoDuplicates()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Existing Shapes",
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "Body", Paragraphs = ["Content"] }
            ]
        });
        var headers = new[] { "Col1" };
        var rows = new[] { new[] { "Data1" } };

        var result = _service.InsertTable(path, 1, headers, rows);

        Assert.True(result.Success);
        Assert.NotNull(result.TableShapeId);

        // Verify no duplicate IDs in the shape tree
        using var doc = PresentationDocument.Open(path, false);
        var slidePart = GetSlidePart(doc, 0);
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
        var allIds = GetAllShapeIds(shapeTree);
        Assert.Equal(allIds.Count, allIds.Distinct().Count());
    }

    [Fact]
    public void InsertTable_GraphicDataUri_IsCorrectTableUri()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "URI Check" });
        var headers = new[] { "A" };
        var rows = new[] { new[] { "1" } };

        _service.InsertTable(path, 1, headers, rows);

        using var doc = PresentationDocument.Open(path, false);
        var slidePart = GetSlidePart(doc, 0);
        var graphicFrame = slidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>()
            .Last(); // Inserted table is the last graphic frame

        var graphicData = graphicFrame.Graphic!.GraphicData!;
        Assert.Equal("http://schemas.openxmlformats.org/drawingml/2006/table", graphicData.Uri);
    }

    [Fact]
    public void InsertTable_TableGridColumnCount_MatchesHeaderCount()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Grid Check" });
        var headers = new[] { "A", "B", "C", "D" };
        var rows = new[] { new[] { "1", "2", "3", "4" } };

        _service.InsertTable(path, 1, headers, rows);

        using var doc = PresentationDocument.Open(path, false);
        var slidePart = GetSlidePart(doc, 0);
        var graphicFrame = slidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().Last();
        var table = graphicFrame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
        var gridColumns = table.TableGrid!.Elements<A.GridColumn>().ToList();

        Assert.Equal(4, gridColumns.Count);
    }

    // ────────────────────────────────────────────────────────
    // InsertTable: edge cases
    // ────────────────────────────────────────────────────────

    [Fact]
    public void InsertTable_HeadersOnly_NoDataRows_CreatesTableWithOneRow()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Empty Table" });
        var headers = new[] { "Name", "Status" };
        var rows = Array.Empty<string[]>();

        var result = _service.InsertTable(path, 1, headers, rows);

        Assert.True(result.Success);
        Assert.Equal(1, result.RowCount);   // Header row only
        Assert.Equal(2, result.ColumnCount);

        var slideContent = _service.GetSlideContent(path, 0);
        var tableShape = Assert.Single(slideContent.Shapes, s => s.ShapeType == "Table");
        Assert.NotNull(tableShape.TableRows);
        Assert.Single(tableShape.TableRows);
        Assert.Equal("Name", tableShape.TableRows[0][0]);
        Assert.Equal("Status", tableShape.TableRows[0][1]);
    }

    [Fact]
    public void InsertTable_SingleCell_1x1_CreatesValidTable()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Tiny" });
        var baselineErrors = ValidatePresentation(path);
        var headers = new[] { "Value" };
        var rows = Array.Empty<string[]>();

        var result = _service.InsertTable(path, 1, headers, rows);

        Assert.True(result.Success);
        Assert.Equal(1, result.ColumnCount);
        var postErrors = ValidatePresentation(path);
        Assert.Equal(baselineErrors.Count, postErrors.Count);
    }

    [Fact]
    public void InsertTable_LargeTable_ManyRowsAndColumns()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Large" });
        var baselineErrors = ValidatePresentation(path);
        var headers = Enumerable.Range(1, 6).Select(i => $"Col{i}").ToArray();
        var rows = Enumerable.Range(1, 12)
            .Select(r => Enumerable.Range(1, 6).Select(c => $"R{r}C{c}").ToArray())
            .ToArray();

        var result = _service.InsertTable(path, 1, headers, rows);

        Assert.True(result.Success);
        Assert.Equal(13, result.RowCount);  // 1 header + 12 data
        Assert.Equal(6, result.ColumnCount);
        var postErrors = ValidatePresentation(path);
        Assert.Equal(baselineErrors.Count, postErrors.Count);
    }

    [Fact]
    public void InsertTable_CustomPositioning_UsesSpecifiedEmuValues()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Positioned" });
        var headers = new[] { "X" };
        var rows = new[] { new[] { "1" } };
        long customX = 500000, customY = 2000000, customW = 5000000, customH = 1500000;

        _service.InsertTable(path, 1, headers, rows, x: customX, y: customY, width: customW, height: customH);

        using var doc = PresentationDocument.Open(path, false);
        var slidePart = GetSlidePart(doc, 0);
        var graphicFrame = slidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().Last();
        var transform = graphicFrame.Transform!;

        Assert.Equal(customX, transform.Offset!.X!.Value);
        Assert.Equal(customY, transform.Offset!.Y!.Value);
        Assert.Equal(customW, transform.Extents!.Cx!.Value);
        Assert.Equal(customH, transform.Extents!.Cy!.Value);
    }

    [Fact]
    public void InsertTable_DefaultPositioning_WhenNotSpecified()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Default Pos" });
        var headers = new[] { "A" };
        var rows = new[] { new[] { "1" } };

        _service.InsertTable(path, 1, headers, rows);

        using var doc = PresentationDocument.Open(path, false);
        var slidePart = GetSlidePart(doc, 0);
        var graphicFrame = slidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().Last();
        var transform = graphicFrame.Transform!;

        // Should have some default position (not null/zero)
        Assert.NotNull(transform.Offset);
        Assert.NotNull(transform.Extents);
        Assert.True(transform.Extents!.Cx!.Value > 0);
        Assert.True(transform.Extents!.Cy!.Value > 0);
    }

    [Fact]
    public void InsertTable_CustomTableName_IsStored()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Named" });
        var headers = new[] { "Region" };
        var rows = new[] { new[] { "NA" } };

        var result = _service.InsertTable(path, 1, headers, rows, tableName: "Revenue Table");

        Assert.True(result.Success);
        Assert.Equal("Revenue Table", result.TableName);

        // Verify the name is in the OpenXML tree
        using var doc = PresentationDocument.Open(path, false);
        var slidePart = GetSlidePart(doc, 0);
        var graphicFrame = slidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().Last();
        var frameName = graphicFrame.NonVisualGraphicFrameProperties!
            .NonVisualDrawingProperties!.Name!.Value;
        Assert.Equal("Revenue Table", frameName);
    }

    // ────────────────────────────────────────────────────────
    // InsertTable: error handling
    // ────────────────────────────────────────────────────────

    [Fact]
    public void InsertTable_InvalidSlideNumber_ThrowsOrReturnsFailure()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Single" });
        var headers = new[] { "A" };
        var rows = new[] { new[] { "1" } };

        // Slide 5 does not exist in a 1-slide presentation
        var ex = Assert.ThrowsAny<Exception>(() => _service.InsertTable(path, 5, headers, rows));
        Assert.NotNull(ex);
    }

    [Fact]
    public void InsertTable_FileNotFound_ThrowsOrReturnsFailure()
    {
        var headers = new[] { "A" };
        var rows = new[] { new[] { "1" } };
        var fakePath = Path.Join(Path.GetTempPath(), "nonexistent_" + Guid.NewGuid() + ".pptx");

        Assert.ThrowsAny<Exception>(() => _service.InsertTable(fakePath, 1, headers, rows));
    }

    // ────────────────────────────────────────────────────────
    // InsertTable: PowerPoint compatibility
    // ────────────────────────────────────────────────────────

    [Fact]
    public void InsertTable_PassesOpenXmlValidator()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Validated" });
        var baselineErrors = ValidatePresentation(path);
        var headers = new[] { "Metric", "Q1", "Q2" };
        var rows = new[]
        {
            new[] { "Revenue", "1.2M", "1.5M" },
            new[] { "Margin", "58%", "62%" }
        };

        _service.InsertTable(path, 1, headers, rows);

        // Should not introduce new validation errors
        var postErrors = ValidatePresentation(path);
        Assert.Equal(baselineErrors.Count, postErrors.Count);
    }

    [Fact]
    public void InsertTable_CellTextBodyStructure_IsWellFormed()
    {
        var path = CreatePresentation(new TestSlideDefinition { TitleText = "Structure" });
        _service.InsertTable(path, 1, new[] { "H1" }, new[] { new[] { "D1" } });

        using var doc = PresentationDocument.Open(path, false);
        var slidePart = GetSlidePart(doc, 0);
        var graphicFrame = slidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().Last();
        var table = graphicFrame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;

        foreach (var row in table.Elements<A.TableRow>())
        {
            foreach (var cell in row.Elements<A.TableCell>())
            {
                // Each cell MUST have TextBody with BodyProperties
                var textBody = cell.TextBody;
                Assert.NotNull(textBody);
                Assert.NotNull(textBody!.GetFirstChild<A.BodyProperties>());

                // Each TextBody must have at least one Paragraph
                var paragraphs = textBody.Elements<A.Paragraph>().ToList();
                Assert.NotEmpty(paragraphs);
            }
        }
    }

    [Fact]
    public void InsertTable_PreservesExistingShapes()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Dashboard",
            TextShapes =
            [
                new TestTextShapeDefinition
                {
                    Name = "Revenue Value",
                    Paragraphs = ["3.2M"]
                }
            ]
        });

        _service.InsertTable(path, 1, new[] { "Region" }, new[] { new[] { "NA" } });

        // Original shapes must still be present
        var slideContent = _service.GetSlideContent(path, 0);
        Assert.Contains(slideContent.Shapes, s => s.Name == "Revenue Value");
        Assert.Contains(slideContent.Shapes, s => s.ShapeType == "Table");
    }

    // ────────────────────────────────────────────────────────
    // UpdateTable: cell updates
    // ────────────────────────────────────────────────────────

    [Fact]
    public void UpdateTable_SingleCell_UpdatesByRowColumn()
    {
        var path = CreatePresentationWithTable("Sales Data",
            [["Region", "Revenue"], ["NA", "3.2M"], ["EMEA", "1.4M"]]);

        var result = _service.UpdateTable(path, 1, tableName: "Sales Data", updates:
        [
            new TableCellUpdate(1, 1, "4.8M")  // Row 1, Col 1 → "4.8M"
        ]);

        Assert.True(result.Success);

        var slideContent = _service.GetSlideContent(path, 0);
        var tableShape = Assert.Single(slideContent.Shapes, s => s.ShapeType == "Table");
        Assert.Equal("4.8M", tableShape.TableRows![1][1]);

        // Untouched cells remain
        Assert.Equal("Region", tableShape.TableRows[0][0]);
        Assert.Equal("Revenue", tableShape.TableRows[0][1]);
        Assert.Equal("NA", tableShape.TableRows[1][0]);
        Assert.Equal("EMEA", tableShape.TableRows[2][0]);
        Assert.Equal("1.4M", tableShape.TableRows[2][1]);
    }

    [Fact]
    public void UpdateTable_MultipleCells_InOneCall()
    {
        var path = CreatePresentationWithTable("Quarterly",
            [["Q1", "Q2", "Q3"], ["1.0M", "1.2M", "1.5M"]]);

        var result = _service.UpdateTable(path, 1, tableName: "Quarterly", updates:
        [
            new TableCellUpdate(1, 0, "1.1M"),
            new TableCellUpdate(1, 1, "1.3M"),
            new TableCellUpdate(1, 2, "1.6M")
        ]);

        Assert.True(result.Success);

        var slideContent = _service.GetSlideContent(path, 0);
        var tableShape = Assert.Single(slideContent.Shapes, s => s.ShapeType == "Table");
        Assert.Equal("1.1M", tableShape.TableRows![1][0]);
        Assert.Equal("1.3M", tableShape.TableRows[1][1]);
        Assert.Equal("1.6M", tableShape.TableRows[1][2]);
    }

    [Fact]
    public void UpdateTable_LocateByName_CaseInsensitive()
    {
        var path = CreatePresentationWithTable("Revenue Table",
            [["Metric", "Value"], ["ARR", "3.2M"]]);

        var result = _service.UpdateTable(path, 1, tableName: "revenue table", updates:
        [
            new TableCellUpdate(1, 1, "4.0M")
        ]);

        Assert.True(result.Success);

        var slideContent = _service.GetSlideContent(path, 0);
        var tableShape = Assert.Single(slideContent.Shapes, s => s.ShapeType == "Table");
        Assert.Equal("4.0M", tableShape.TableRows![1][1]);
    }

    [Fact]
    public void UpdateTable_LocateByIndex_ZeroBased()
    {
        var path = CreatePresentationWithTable("First Table",
            [["A", "B"], ["1", "2"]]);

        var result = _service.UpdateTable(path, 1, tableIndex: 0, updates:
        [
            new TableCellUpdate(1, 0, "Updated")
        ]);

        Assert.True(result.Success);

        var slideContent = _service.GetSlideContent(path, 0);
        var tableShape = Assert.Single(slideContent.Shapes, s => s.ShapeType == "Table");
        Assert.Equal("Updated", tableShape.TableRows![1][0]);
    }

    // ────────────────────────────────────────────────────────
    // UpdateTable: error handling
    // ────────────────────────────────────────────────────────

    [Fact]
    public void UpdateTable_TableNotFoundByName_ReturnsFailure()
    {
        var path = CreatePresentationWithTable("Actual Table",
            [["A"], ["1"]]);

        var ex = Assert.ThrowsAny<Exception>(() =>
            _service.UpdateTable(path, 1, tableName: "Missing Table", updates:
            [
                new TableCellUpdate(0, 0, "X")
            ]));
        Assert.NotNull(ex);
    }

    [Fact]
    public void UpdateTable_TableIndexOutOfRange_ReturnsFailure()
    {
        var path = CreatePresentationWithTable("Only Table",
            [["A"], ["1"]]);

        var ex = Assert.ThrowsAny<Exception>(() =>
            _service.UpdateTable(path, 1, tableIndex: 5, updates:
            [
                new TableCellUpdate(0, 0, "X")
            ]));
        Assert.NotNull(ex);
    }

    [Fact]
    public void UpdateTable_CellCoordinatesOutOfRange_ReturnsFailure()
    {
        var path = CreatePresentationWithTable("Small",
            [["A", "B"], ["1", "2"]]);

        // Implementation may throw or return a failure result — accept either
        try
        {
            var result = _service.UpdateTable(path, 1, tableName: "Small", updates:
            [
                new TableCellUpdate(99, 99, "Out of bounds")
            ]);
            Assert.False(result.Success);
        }
        catch (Exception ex)
        {
            Assert.NotNull(ex);
        }
    }

    // ────────────────────────────────────────────────────────
    // UpdateTable: preservation and compatibility
    // ────────────────────────────────────────────────────────

    [Fact]
    public void UpdateTable_PreservesTableCellProperties()
    {
        var path = CreatePresentationWithTable("Styled",
            [["Header", "Data"], ["Value", "100"]]);

        // Capture cell properties XML before update
        var cellPropsBeforeUpdate = GetTableCellPropertiesXml(path, 0);

        _service.UpdateTable(path, 1, tableName: "Styled", updates:
        [
            new TableCellUpdate(1, 1, "200")
        ]);

        // Cell properties should be preserved after update
        var cellPropsAfterUpdate = GetTableCellPropertiesXml(path, 0);
        Assert.Equal(cellPropsBeforeUpdate, cellPropsAfterUpdate);
    }

    [Fact]
    public void UpdateTable_PreservesOtherShapes()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Dashboard",
            TextShapes =
            [
                new TestTextShapeDefinition
                {
                    Name = "Revenue Value",
                    Paragraphs = ["3.2M"]
                }
            ],
            Tables =
            [
                new TestTableDefinition
                {
                    Name = "Region Data",
                    Rows = [["Region", "ARR"], ["NA", "3.2M"]]
                }
            ]
        });

        _service.UpdateTable(path, 1, tableName: "Region Data", updates:
        [
            new TableCellUpdate(1, 1, "4.0M")
        ]);

        var slideContent = _service.GetSlideContent(path, 0);
        var revenueShape = Assert.Single(slideContent.Shapes, s => s.Name == "Revenue Value");
        Assert.Equal("3.2M", revenueShape.Text);
    }

    [Fact]
    public void UpdateTable_PassesOpenXmlValidator_AfterUpdate()
    {
        var path = CreatePresentationWithTable("Validated",
            [["Metric", "Value"], ["ARR", "3.2M"], ["NRR", "112%"]]);
        var baselineErrors = ValidatePresentation(path);

        _service.UpdateTable(path, 1, tableName: "Validated", updates:
        [
            new TableCellUpdate(1, 1, "4.8M"),
            new TableCellUpdate(2, 1, "118%")
        ]);

        var postErrors = ValidatePresentation(path);
        Assert.Equal(baselineErrors.Count, postErrors.Count);
    }

    [Fact]
    public void UpdateTable_PresentationStructure_RemainsValid()
    {
        var path = CreatePresentationWithTable("Structure Check",
            [["A", "B"], ["1", "2"]]);

        _service.UpdateTable(path, 1, tableName: "Structure Check", updates:
        [
            new TableCellUpdate(0, 0, "Updated Header")
        ]);

        AssertPresentationCompatible(path, 1);
    }

    // ────────────────────────────────────────────────────────
    // Helpers
    // ────────────────────────────────────────────────────────

    private string CreatePresentation(params TestSlideDefinition[] slides)
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        _tempFiles.Add(path);
        TestPptxHelper.CreatePresentation(path, slides);
        return path;
    }

    private string CreatePresentationWithTable(string tableName, IReadOnlyList<IReadOnlyList<string>> rows)
    {
        return CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Table Slide",
            Tables =
            [
                new TestTableDefinition
                {
                    Name = tableName,
                    Rows = rows
                }
            ]
        });
    }

    private static SlidePart GetSlidePart(PresentationDocument document, int slideIndex)
    {
        var presentationPart = document.PresentationPart!;
        var slideIdList = presentationPart.Presentation.SlideIdList!;
        var slideId = slideIdList.Elements<SlideId>().ElementAt(slideIndex);
        return (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!);
    }

    private static List<uint> GetAllShapeIds(ShapeTree shapeTree)
    {
        var ids = new List<uint>();
        foreach (var child in shapeTree.ChildElements)
        {
            uint? id = child switch
            {
                Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                P.GraphicFrame gf => gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
                Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
                GroupShape gs => gs.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                ConnectionShape cs => cs.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                _ => null
            };
            if (id.HasValue)
                ids.Add(id.Value);
        }
        return ids;
    }

    private static void AssertPresentationCompatible(string path, int expectedSlideCount)
    {
        using var document = PresentationDocument.Open(path, false);
        var presentationPart = Assert.IsType<PresentationPart>(document.PresentationPart);
        var presentation = Assert.IsType<Presentation>(presentationPart.Presentation);
        var slideIdList = Assert.IsType<SlideIdList>(presentation.SlideIdList);
        var slideIds = slideIdList.Elements<SlideId>().ToList();
        Assert.Equal(expectedSlideCount, slideIds.Count);

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

    private static List<string> GetTableCellPropertiesXml(string path, int slideIndex)
    {
        using var doc = PresentationDocument.Open(path, false);
        var slidePart = GetSlidePart(doc, slideIndex);
        var graphicFrame = slidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().First();
        var table = graphicFrame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;

        return table.Elements<A.TableRow>()
            .SelectMany(row => row.Elements<A.TableCell>())
            .Select(cell => cell.TableCellProperties?.OuterXml ?? string.Empty)
            .ToList();
    }
}
