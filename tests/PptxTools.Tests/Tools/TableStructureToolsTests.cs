using System.Text.Json;

namespace PptxTools.Tests.Tools;

/// <summary>
/// Tool-level tests for pptx_table_structure MCP tool.
/// Written for Issue #135 — extended table operations.
/// Verifies JSON serialization, error handling, and parameter validation at the MCP tool layer.
/// </summary>
[Trait("Category", "Integration")]
public class TableStructureToolsTests : PptxTestBase
{
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public TableStructureToolsTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    // ────────────────────────────────────────────────────────
    // JSON result structure per action
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task AddRow_ReturnsStructuredJsonWithSuccess()
    {
        var path = CreatePptxWithTable("Tool Add", [["H1", "H2"], ["A", "B"]]);

        var json = await _tools.pptx_table_structure(path, 1,
            TableStructureAction.AddRow, tableName: "Tool Add", cellValues: ["C", "D"]);

        var result = JsonSerializer.Deserialize<TableStructureResult>(json);
        Assert.NotNull(result);
        Assert.True(result.Success);
        Assert.Equal(1, result.SlideNumber);
        Assert.Equal("AddRow", result.Action);
        Assert.Equal("Tool Add", result.TableName);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
    }

    [Fact]
    public async Task DeleteRow_ReturnsStructuredJsonWithSuccess()
    {
        var path = CreatePptxWithTable("Tool Del", [["H"], ["R1"], ["R2"]]);

        var json = await _tools.pptx_table_structure(path, 1,
            TableStructureAction.DeleteRow, tableName: "Tool Del", rowIndex: 1);

        var result = JsonSerializer.Deserialize<TableStructureResult>(json);
        Assert.NotNull(result);
        Assert.True(result.Success);
        Assert.Equal("DeleteRow", result.Action);
        Assert.Equal(2, result.RowCount);
    }

    [Fact]
    public async Task AddColumn_ReturnsStructuredJsonWithSuccess()
    {
        var path = CreatePptxWithTable("Tool Col", [["A"], ["1"]]);

        var json = await _tools.pptx_table_structure(path, 1,
            TableStructureAction.AddColumn, tableName: "Tool Col", headerText: "B");

        var result = JsonSerializer.Deserialize<TableStructureResult>(json);
        Assert.NotNull(result);
        Assert.True(result.Success);
        Assert.Equal("AddColumn", result.Action);
        Assert.Equal(2, result.ColumnCount);
    }

    [Fact]
    public async Task DeleteColumn_ReturnsStructuredJsonWithSuccess()
    {
        var path = CreatePptxWithTable("Tool DC", [["A", "B"], ["1", "2"]]);

        var json = await _tools.pptx_table_structure(path, 1,
            TableStructureAction.DeleteColumn, tableName: "Tool DC", columnIndex: 0);

        var result = JsonSerializer.Deserialize<TableStructureResult>(json);
        Assert.NotNull(result);
        Assert.True(result.Success);
        Assert.Equal("DeleteColumn", result.Action);
        Assert.Equal(1, result.ColumnCount);
    }

    [Fact]
    public async Task MergeCells_ReturnsStructuredJsonWithSuccess()
    {
        var path = CreatePptxWithTable("Tool Merge", [["A", "B"], ["C", "D"]]);

        var json = await _tools.pptx_table_structure(path, 1,
            TableStructureAction.MergeCells, tableName: "Tool Merge",
            startRow: 0, startCol: 0, endRow: 1, endCol: 1);

        var result = JsonSerializer.Deserialize<TableStructureResult>(json);
        Assert.NotNull(result);
        Assert.True(result.Success);
        Assert.Equal("MergeCells", result.Action);
    }

    // ────────────────────────────────────────────────────────
    // File not found — all actions
    // ────────────────────────────────────────────────────────

    [Theory]
    [InlineData(TableStructureAction.AddRow)]
    [InlineData(TableStructureAction.DeleteRow)]
    [InlineData(TableStructureAction.AddColumn)]
    [InlineData(TableStructureAction.DeleteColumn)]
    [InlineData(TableStructureAction.MergeCells)]
    public async Task FileNotFound_ReturnsFailureJson(TableStructureAction action)
    {
        var fakePath = "C:\\does-not-exist\\file.pptx";

        var json = await _tools.pptx_table_structure(fakePath, 1, action,
            rowIndex: 0, columnIndex: 0,
            startRow: 0, startCol: 0, endRow: 0, endCol: 0);

        using var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.GetProperty("Success").GetBoolean());
        Assert.Contains("not found",
            doc.RootElement.GetProperty("Message").GetString()!,
            StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────
    // Null/empty file path
    // ────────────────────────────────────────────────────────

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public async Task EmptyFilePath_ReturnsFailure(string emptyPath)
    {
        var json = await _tools.pptx_table_structure(emptyPath, 1,
            TableStructureAction.AddRow);

        using var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.GetProperty("Success").GetBoolean());
    }

    // ────────────────────────────────────────────────────────
    // Missing required parameters
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task DeleteRow_WithoutRowIndex_ReturnsFailure()
    {
        var path = CreatePptxWithTable("No Idx", [["H"], ["R1"], ["R2"]]);

        var json = await _tools.pptx_table_structure(path, 1,
            TableStructureAction.DeleteRow, tableName: "No Idx");

        using var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.GetProperty("Success").GetBoolean());
        Assert.Contains("rowIndex",
            doc.RootElement.GetProperty("Message").GetString()!,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteColumn_WithoutColumnIndex_ReturnsFailure()
    {
        var path = CreatePptxWithTable("No Col Idx", [["A", "B"], ["1", "2"]]);

        var json = await _tools.pptx_table_structure(path, 1,
            TableStructureAction.DeleteColumn, tableName: "No Col Idx");

        using var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.GetProperty("Success").GetBoolean());
        Assert.Contains("columnIndex",
            doc.RootElement.GetProperty("Message").GetString()!,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task MergeCells_WithoutAllCoordinates_ReturnsFailure()
    {
        var path = CreatePptxWithTable("No Coords", [["A", "B"], ["C", "D"]]);

        var json = await _tools.pptx_table_structure(path, 1,
            TableStructureAction.MergeCells, tableName: "No Coords",
            startRow: 0, startCol: 0);
        // endRow and endCol not provided

        using var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.GetProperty("Success").GetBoolean());
    }

    // ────────────────────────────────────────────────────────
    // Service error surfaced as structured JSON failure
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task DeleteRow_LastRow_ReturnsFailureJson()
    {
        var path = CreatePptxWithTable("Solo Row", [["Only"]]);

        var json = await _tools.pptx_table_structure(path, 1,
            TableStructureAction.DeleteRow, tableName: "Solo Row", rowIndex: 0);

        using var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.GetProperty("Success").GetBoolean());
        Assert.Contains("last remaining",
            doc.RootElement.GetProperty("Message").GetString()!,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteColumn_LastColumn_ReturnsFailureJson()
    {
        var path = CreatePptxWithTable("Solo Col", [["Only"], ["Data"]]);

        var json = await _tools.pptx_table_structure(path, 1,
            TableStructureAction.DeleteColumn, tableName: "Solo Col", columnIndex: 0);

        using var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.GetProperty("Success").GetBoolean());
        Assert.Contains("last remaining",
            doc.RootElement.GetProperty("Message").GetString()!,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task InvalidSlideNumber_ReturnsFailureJson()
    {
        var path = CreatePptxWithTable("Bad Slide", [["A"], ["1"]]);

        var json = await _tools.pptx_table_structure(path, 99,
            TableStructureAction.AddRow, tableName: "Bad Slide");

        using var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.GetProperty("Success").GetBoolean());
    }

    // ────────────────────────────────────────────────────────
    // Helpers
    // ────────────────────────────────────────────────────────

    private string CreatePptxWithTable(string tableName, IReadOnlyList<IReadOnlyList<string>> rows)
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
}
