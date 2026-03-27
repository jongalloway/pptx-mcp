using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests.Services;

/// <summary>
/// Service-level tests for table structural operations: AddTableRow, DeleteTableRow,
/// AddTableColumn, DeleteTableColumn, MergeTableCells.
/// Written for Issue #135 — extended table operations.
/// </summary>
[Trait("Category", "Unit")]
public class TableStructureTests : PptxTestBase
{
    // ────────────────────────────────────────────────────────
    // AddTableRow
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AddTableRow_AtEnd_AppendsRowAndUpdatesCount()
    {
        var path = CreateTable("Add End", [["H1", "H2"], ["A", "B"]]);

        var result = Service.AddTableRow(path, 1, tableName: "Add End");

        Assert.True(result.Success);
        Assert.Equal("AddRow", result.Action);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
    }

    [Fact]
    public void AddTableRow_AtSpecificIndex_InsertsAtCorrectPosition()
    {
        var path = CreateTable("Insert Mid", [["H1"], ["R0"], ["R1"]]);

        var result = Service.AddTableRow(path, 1, rowIndex: 1, cellValues: ["New"], tableName: "Insert Mid");

        Assert.True(result.Success);
        Assert.Equal(4, result.RowCount);

        // Verify the inserted row is at index 1
        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal("New", table.TableRows![1][0]);
        Assert.Equal("R0", table.TableRows[2][0]);
    }

    [Fact]
    public void AddTableRow_WithCellValues_PopulatesCells()
    {
        var path = CreateTable("Values", [["Name", "Score"], ["Alice", "95"]]);

        Service.AddTableRow(path, 1, cellValues: ["Bob", "88"], tableName: "Values");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal("Bob", table.TableRows![2][0]);
        Assert.Equal("88", table.TableRows[2][1]);
    }

    [Fact]
    public void AddTableRow_WithFewerValues_PadsWithEmpty()
    {
        var path = CreateTable("Pad", [["A", "B", "C"], ["1", "2", "3"]]);

        Service.AddTableRow(path, 1, cellValues: ["Only One"], tableName: "Pad");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal("Only One", table.TableRows![2][0]);
        Assert.Equal("", table.TableRows[2][1]);
        Assert.Equal("", table.TableRows[2][2]);
    }

    [Fact]
    public void AddTableRow_AtIndexZero_InsertsBeforeFirstRow()
    {
        var path = CreateTable("Prepend", [["H1"], ["Original"]]);

        Service.AddTableRow(path, 1, rowIndex: 0, cellValues: ["Before Header"], tableName: "Prepend");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal(3, table.TableRows!.Count);
        Assert.Equal("Before Header", table.TableRows[0][0]);
        Assert.Equal("H1", table.TableRows[1][0]);
    }

    // ────────────────────────────────────────────────────────
    // DeleteTableRow
    // ────────────────────────────────────────────────────────

    [Fact]
    public void DeleteTableRow_FirstRow_RemovesAndPreservesOthers()
    {
        var path = CreateTable("Del First", [["H1"], ["R1"], ["R2"]]);

        var result = Service.DeleteTableRow(path, 1, rowIndex: 0, tableName: "Del First");

        Assert.True(result.Success);
        Assert.Equal("DeleteRow", result.Action);
        Assert.Equal(2, result.RowCount);

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal("R1", table.TableRows![0][0]);
        Assert.Equal("R2", table.TableRows[1][0]);
    }

    [Fact]
    public void DeleteTableRow_LastRow_RemovesAndPreservesOthers()
    {
        var path = CreateTable("Del Last", [["H1"], ["R1"], ["R2"]]);

        var result = Service.DeleteTableRow(path, 1, rowIndex: 2, tableName: "Del Last");

        Assert.True(result.Success);
        Assert.Equal(2, result.RowCount);

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal("H1", table.TableRows![0][0]);
        Assert.Equal("R1", table.TableRows[1][0]);
    }

    [Fact]
    public void DeleteTableRow_MiddleRow_RemovesCorrectRow()
    {
        var path = CreateTable("Del Mid", [["H"], ["A"], ["B"], ["C"]]);

        Service.DeleteTableRow(path, 1, rowIndex: 2, tableName: "Del Mid");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal(3, table.TableRows!.Count);
        Assert.Equal("H", table.TableRows[0][0]);
        Assert.Equal("A", table.TableRows[1][0]);
        Assert.Equal("C", table.TableRows[2][0]);
    }

    [Fact]
    public void DeleteTableRow_FromSingleRowTable_ThrowsInvalidOperation()
    {
        var path = CreateTable("Solo", [["Only Row"]]);

        Assert.Throws<InvalidOperationException>(() =>
            Service.DeleteTableRow(path, 1, rowIndex: 0, tableName: "Solo"));
    }

    // ────────────────────────────────────────────────────────
    // AddTableColumn
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AddTableColumn_AtEnd_AppendsColumnToAllRows()
    {
        var path = CreateTable("Add Col", [["H1", "H2"], ["A", "B"]]);

        var result = Service.AddTableColumn(path, 1, headerText: "H3", tableName: "Add Col");

        Assert.True(result.Success);
        Assert.Equal("AddColumn", result.Action);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(2, result.RowCount);

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal("H3", table.TableRows![0][2]);
        Assert.Equal("", table.TableRows[1][2]);
    }

    [Fact]
    public void AddTableColumn_AtSpecificIndex_InsertsCorrectly()
    {
        var path = CreateTable("Col Mid", [["A", "C"], ["1", "3"]]);

        Service.AddTableColumn(path, 1, columnIndex: 1, headerText: "B", tableName: "Col Mid");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal("A", table.TableRows![0][0]);
        Assert.Equal("B", table.TableRows[0][1]);
        Assert.Equal("C", table.TableRows[0][2]);
    }

    [Fact]
    public void AddTableColumn_AllRowsGetNewCell()
    {
        var path = CreateTable("All Rows", [["H"], ["R1"], ["R2"], ["R3"]]);

        Service.AddTableColumn(path, 1, tableName: "All Rows");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        foreach (var row in table.TableRows!)
            Assert.Equal(2, row.Count);
    }

    [Fact]
    public void AddTableColumn_UpdatesTableGrid()
    {
        var path = CreateTable("Grid Check", [["A"], ["1"]]);

        Service.AddTableColumn(path, 1, tableName: "Grid Check");

        using var doc = PresentationDocument.Open(path, false);
        var slidePart = GetSlidePart(doc, 0);
        var graphicFrame = slidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().First();
        var table = graphicFrame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
        var gridColumns = table.TableGrid!.Elements<A.GridColumn>().Count();
        Assert.Equal(2, gridColumns);
    }

    // ────────────────────────────────────────────────────────
    // DeleteTableColumn
    // ────────────────────────────────────────────────────────

    [Fact]
    public void DeleteTableColumn_FirstColumn_RemovesFromAllRows()
    {
        var path = CreateTable("Del Col First", [["A", "B", "C"], ["1", "2", "3"]]);

        var result = Service.DeleteTableColumn(path, 1, columnIndex: 0, tableName: "Del Col First");

        Assert.True(result.Success);
        Assert.Equal("DeleteColumn", result.Action);
        Assert.Equal(2, result.ColumnCount);

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal("B", table.TableRows![0][0]);
        Assert.Equal("C", table.TableRows[0][1]);
    }

    [Fact]
    public void DeleteTableColumn_LastColumn_RemovesCorrectly()
    {
        var path = CreateTable("Del Col Last", [["A", "B", "C"], ["1", "2", "3"]]);

        Service.DeleteTableColumn(path, 1, columnIndex: 2, tableName: "Del Col Last");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal(2, table.TableRows![0].Count);
        Assert.Equal("A", table.TableRows[0][0]);
        Assert.Equal("B", table.TableRows[0][1]);
    }

    [Fact]
    public void DeleteTableColumn_MiddleColumn_RemovesAndPreservesOthers()
    {
        var path = CreateTable("Del Col Mid", [["A", "B", "C"], ["1", "2", "3"]]);

        Service.DeleteTableColumn(path, 1, columnIndex: 1, tableName: "Del Col Mid");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal("A", table.TableRows![0][0]);
        Assert.Equal("C", table.TableRows[0][1]);
        Assert.Equal("1", table.TableRows[1][0]);
        Assert.Equal("3", table.TableRows[1][1]);
    }

    [Fact]
    public void DeleteTableColumn_FromSingleColumnTable_ThrowsInvalidOperation()
    {
        var path = CreateTable("Sole Col", [["Only"], ["Data"]]);

        Assert.Throws<InvalidOperationException>(() =>
            Service.DeleteTableColumn(path, 1, columnIndex: 0, tableName: "Sole Col"));
    }

    [Fact]
    public void DeleteTableColumn_UpdatesTableGrid()
    {
        var path = CreateTable("Grid Del", [["A", "B", "C"], ["1", "2", "3"]]);

        Service.DeleteTableColumn(path, 1, columnIndex: 1, tableName: "Grid Del");

        using var doc = PresentationDocument.Open(path, false);
        var slidePart = GetSlidePart(doc, 0);
        var graphicFrame = slidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().First();
        var table = graphicFrame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
        var gridColumns = table.TableGrid!.Elements<A.GridColumn>().Count();
        Assert.Equal(2, gridColumns);
    }

    // ────────────────────────────────────────────────────────
    // MergeCells
    // ────────────────────────────────────────────────────────

    [Fact]
    public void MergeCells_2x2Range_SetsGridSpanAndRowSpan()
    {
        var path = CreateTable("Merge 2x2",
            [["A", "B", "C"], ["D", "E", "F"], ["G", "H", "I"]]);

        var result = Service.MergeTableCells(path, 1,
            startRow: 0, startCol: 0, endRow: 1, endCol: 1, tableName: "Merge 2x2");

        Assert.True(result.Success);
        Assert.Equal("MergeCells", result.Action);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(3, result.ColumnCount);

        // Verify OpenXML merge attributes (set on A.TableCell, not TableCellProperties)
        using var doc = PresentationDocument.Open(path, false);
        var table = GetOpenXmlTable(doc, 0);
        var rows = table.Elements<A.TableRow>().ToList();

        // Anchor cell [0,0]: GridSpan=2, RowSpan=2
        var anchor = rows[0].Elements<A.TableCell>().ElementAt(0);
        Assert.Equal(2, anchor.GridSpan?.Value);
        Assert.Equal(2, anchor.RowSpan?.Value);

        // [0,1]: HorizontalMerge=true
        var rightOfAnchor = rows[0].Elements<A.TableCell>().ElementAt(1);
        Assert.True(rightOfAnchor.HorizontalMerge?.Value);

        // [1,0]: VerticalMerge=true
        var belowAnchor = rows[1].Elements<A.TableCell>().ElementAt(0);
        Assert.True(belowAnchor.VerticalMerge?.Value);

        // [1,1]: Both HorizontalMerge and VerticalMerge
        var diagonalCell = rows[1].Elements<A.TableCell>().ElementAt(1);
        Assert.True(diagonalCell.HorizontalMerge?.Value);
        Assert.True(diagonalCell.VerticalMerge?.Value);
    }

    [Fact]
    public void MergeCells_EntireRow_SetsGridSpanOnly()
    {
        var path = CreateTable("Merge Row", [["A", "B", "C"], ["D", "E", "F"]]);

        var result = Service.MergeTableCells(path, 1,
            startRow: 0, startCol: 0, endRow: 0, endCol: 2, tableName: "Merge Row");

        Assert.True(result.Success);

        using var doc = PresentationDocument.Open(path, false);
        var table = GetOpenXmlTable(doc, 0);
        var firstRow = table.Elements<A.TableRow>().First();
        var anchor = firstRow.Elements<A.TableCell>().First();
        Assert.Equal(3, anchor.GridSpan?.Value);
        Assert.Null(anchor.RowSpan); // single row merge has no RowSpan
    }

    [Fact]
    public void MergeCells_EntireColumn_SetsRowSpanOnly()
    {
        var path = CreateTable("Merge Col", [["A", "B"], ["C", "D"], ["E", "F"]]);

        var result = Service.MergeTableCells(path, 1,
            startRow: 0, startCol: 0, endRow: 2, endCol: 0, tableName: "Merge Col");

        Assert.True(result.Success);

        using var doc = PresentationDocument.Open(path, false);
        var table = GetOpenXmlTable(doc, 0);
        var anchor = table.Elements<A.TableRow>().First()
            .Elements<A.TableCell>().First();
        Assert.Equal(3, anchor.RowSpan?.Value);
        Assert.Null(anchor.GridSpan); // single column merge has no GridSpan
    }

    [Theory]
    [InlineData(-1, 0, 1, 1, "startRow")]
    [InlineData(0, -1, 1, 1, "startCol")]
    [InlineData(0, 0, 99, 1, "endRow")]
    [InlineData(0, 0, 1, 99, "endCol")]
    public void MergeCells_InvalidRange_ThrowsArgumentOutOfRange(
        int startRow, int startCol, int endRow, int endCol, string expectedParam)
    {
        var path = CreateTable("Bad Range", [["A", "B"], ["C", "D"], ["E", "F"]]);

        var ex = Assert.Throws<ArgumentOutOfRangeException>(() =>
            Service.MergeTableCells(path, 1,
                startRow, startCol, endRow, endCol, tableName: "Bad Range"));
        Assert.Contains(expectedParam, ex.ParamName!, StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────
    // Edge cases: slide/table identification
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AddTableRow_WrongSlideNumber_Throws()
    {
        var path = CreateTable("Edge", [["A"], ["1"]]);

        Assert.ThrowsAny<Exception>(() =>
            Service.AddTableRow(path, 99, tableName: "Edge"));
    }

    [Fact]
    public void DeleteTableRow_NonExistentTableName_Throws()
    {
        var path = CreateTable("Real", [["A"], ["1"], ["2"]]);

        Assert.Throws<InvalidOperationException>(() =>
            Service.DeleteTableRow(path, 1, rowIndex: 0, tableName: "Ghost"));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(99)]
    public void AddTableRow_NegativeOrPastEndIndex_Throws(int badIndex)
    {
        var path = CreateTable("Bounds", [["A"], ["1"]]);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            Service.AddTableRow(path, 1, rowIndex: badIndex, tableName: "Bounds"));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(5)]
    public void DeleteTableRow_OutOfRangeIndex_Throws(int badIndex)
    {
        var path = CreateTable("OOB", [["H"], ["R1"], ["R2"]]);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            Service.DeleteTableRow(path, 1, rowIndex: badIndex, tableName: "OOB"));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(99)]
    public void AddTableColumn_OutOfRangeIndex_Throws(int badIndex)
    {
        var path = CreateTable("Col OOB", [["A", "B"], ["1", "2"]]);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            Service.AddTableColumn(path, 1, columnIndex: badIndex, tableName: "Col OOB"));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(5)]
    public void DeleteTableColumn_OutOfRangeIndex_Throws(int badIndex)
    {
        var path = CreateTable("Del OOB", [["A", "B"], ["1", "2"]]);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            Service.DeleteTableColumn(path, 1, columnIndex: badIndex, tableName: "Del OOB"));
    }

    [Fact]
    public void AddTableColumn_FileNotFound_Throws()
    {
        var fakePath = Path.Join(Path.GetTempPath(), "nonexistent_" + Guid.NewGuid() + ".pptx");

        Assert.ThrowsAny<Exception>(() =>
            Service.AddTableColumn(fakePath, 1));
    }

    [Fact]
    public void MergeTableCells_NoTableOnSlide_Throws()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition { TitleText = "No Table" });

        Assert.Throws<InvalidOperationException>(() =>
            Service.MergeTableCells(path, 1, 0, 0, 0, 0));
    }

    // ────────────────────────────────────────────────────────
    // Table identification: by index
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AddTableRow_ByTableIndex_TargetsCorrectTable()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Multi Table",
            Tables =
            [
                new TestTableDefinition { Name = "First", Rows = [["A"], ["1"]] },
                new TestTableDefinition { Name = "Second", Rows = [["X", "Y"], ["a", "b"]] }
            ]
        });

        var result = Service.AddTableRow(path, 1, tableIndex: 1, cellValues: ["c", "d"]);

        Assert.True(result.Success);
        Assert.Equal("Second", result.TableName);
        Assert.Equal(3, result.RowCount);
    }

    // ────────────────────────────────────────────────────────
    // Data preservation after structural changes
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AddTableRow_PreservesExistingCellText()
    {
        var path = CreateTable("Preserve", [["Name", "Score"], ["Alice", "95"]]);

        Service.AddTableRow(path, 1, tableName: "Preserve");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal("Name", table.TableRows![0][0]);
        Assert.Equal("Score", table.TableRows[0][1]);
        Assert.Equal("Alice", table.TableRows[1][0]);
        Assert.Equal("95", table.TableRows[1][1]);
    }

    [Fact]
    public void DeleteTableColumn_PreservesRemainingCellText()
    {
        var path = CreateTable("Keep", [["A", "B", "C"], ["1", "2", "3"]]);

        Service.DeleteTableColumn(path, 1, columnIndex: 1, tableName: "Keep");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal("A", table.TableRows![0][0]);
        Assert.Equal("C", table.TableRows[0][1]);
        Assert.Equal("1", table.TableRows[1][0]);
        Assert.Equal("3", table.TableRows[1][1]);
    }

    [Fact]
    public void DeleteTableRow_PreservesRemainingCellText()
    {
        var path = CreateTable("KeepRows", [["H1", "H2"], ["A", "B"], ["C", "D"]]);

        Service.DeleteTableRow(path, 1, rowIndex: 1, tableName: "KeepRows");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal(2, table.TableRows!.Count);
        Assert.Equal("H1", table.TableRows[0][0]);
        Assert.Equal("H2", table.TableRows[0][1]);
        Assert.Equal("C", table.TableRows[1][0]);
        Assert.Equal("D", table.TableRows[1][1]);
    }

    // ────────────────────────────────────────────────────────
    // Roundtrip: structural change + GetSlideContent
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AddColumn_ThenGetSlideContent_ReturnsUpdatedTable()
    {
        var path = CreateTable("Roundtrip", [["Region"], ["NA"], ["EMEA"]]);

        Service.AddTableColumn(path, 1, headerText: "Revenue", tableName: "Roundtrip");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal(2, table.TableRows![0].Count);
        Assert.Equal("Revenue", table.TableRows[0][1]);
    }

    [Fact]
    public void DeleteRow_ThenAddRow_RoundTripsCorrectly()
    {
        var path = CreateTable("Combo", [["H"], ["Old1"], ["Old2"]]);

        Service.DeleteTableRow(path, 1, rowIndex: 1, tableName: "Combo");
        Service.AddTableRow(path, 1, cellValues: ["New"], tableName: "Combo");

        var content = Service.GetSlideContent(path, 0);
        var table = GetTableShape(content);
        Assert.Equal(3, table.TableRows!.Count);
        Assert.Equal("H", table.TableRows[0][0]);
        Assert.Equal("Old2", table.TableRows[1][0]);
        Assert.Equal("New", table.TableRows[2][0]);
    }

    // ────────────────────────────────────────────────────────
    // OpenXML validation
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AddTableRow_PassesOpenXmlValidator()
    {
        var path = CreateTable("Valid Add", [["A", "B"], ["1", "2"]]);
        var baseline = ValidatePresentation(path);

        Service.AddTableRow(path, 1, cellValues: ["3", "4"], tableName: "Valid Add");

        var postErrors = ValidatePresentation(path);
        Assert.Equal(baseline.Count, postErrors.Count);
    }

    [Fact]
    public void DeleteTableColumn_PassesOpenXmlValidator()
    {
        var path = CreateTable("Valid Del", [["A", "B", "C"], ["1", "2", "3"]]);
        var baseline = ValidatePresentation(path);

        Service.DeleteTableColumn(path, 1, columnIndex: 1, tableName: "Valid Del");

        var postErrors = ValidatePresentation(path);
        Assert.Equal(baseline.Count, postErrors.Count);
    }

    [Fact]
    public void MergeTableCells_PassesOpenXmlValidator()
    {
        var path = CreateTable("Valid Merge", [["A", "B"], ["C", "D"]]);
        var baseline = ValidatePresentation(path);

        Service.MergeTableCells(path, 1, 0, 0, 1, 1, tableName: "Valid Merge");

        var postErrors = ValidatePresentation(path);
        Assert.Equal(baseline.Count, postErrors.Count);
    }

    // ────────────────────────────────────────────────────────
    // Result metadata
    // ────────────────────────────────────────────────────────

    [Fact]
    public void AddTableRow_Result_ContainsSlideNumberAndTableName()
    {
        var path = CreateTable("Meta", [["H"], ["R"]]);

        var result = Service.AddTableRow(path, 1, tableName: "Meta");

        Assert.Equal(1, result.SlideNumber);
        Assert.Equal("Meta", result.TableName);
        Assert.Contains("Meta", result.Message);
    }

    [Fact]
    public void DeleteTableColumn_Result_ReportsUpdatedColumnCount()
    {
        var path = CreateTable("Count", [["A", "B", "C"], ["1", "2", "3"]]);

        var result = Service.DeleteTableColumn(path, 1, columnIndex: 0, tableName: "Count");

        Assert.Equal(2, result.ColumnCount);
        Assert.Contains("2", result.Message);
    }

    // ────────────────────────────────────────────────────────
    // Helpers
    // ────────────────────────────────────────────────────────

    private string CreateTable(string tableName, IReadOnlyList<IReadOnlyList<string>> rows)
    {
        return CreatePptxWithSlides(new TestSlideDefinition
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

    private static ShapeContent GetTableShape(SlideContent content)
    {
        return Assert.Single(content.Shapes, s => s.ShapeType == "Table");
    }

    private static SlidePart GetSlidePart(PresentationDocument document, int slideIndex)
    {
        var presentationPart = document.PresentationPart!;
        var slideIdList = presentationPart.Presentation.SlideIdList!;
        var slideId = slideIdList.Elements<SlideId>().ElementAt(slideIndex);
        return (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!);
    }

    private static A.Table GetOpenXmlTable(PresentationDocument document, int slideIndex)
    {
        var slidePart = GetSlidePart(document, slideIndex);
        var graphicFrame = slidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().First();
        return graphicFrame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
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
