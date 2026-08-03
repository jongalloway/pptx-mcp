using System.Text.Json;

namespace PptxTools.Tests.Tools;

/// <summary>
/// Tool-level tests for pptx_insert_table and pptx_update_table MCP tools.
/// Written proactively for Issue #36 — table insert and update tools.
/// These tests verify JSON output format, error handling, and parameter validation at the MCP tool layer.
/// </summary>
[Trait("Category", "Integration")]
public class TableToolsTests : PptxTestBase
{
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public TableToolsTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    // ────────────────────────────────────────────────────────
    // pptx_insert_table: JSON output
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task pptx_insert_table_ReturnsStructuredJson()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition { TitleText = "Data Slide" });
        var headers = new[] { "Region", "Revenue" };
        var rows = new[] { new[] { "NA", "3.2M" }, new[] { "EMEA", "1.4M" } };

        var result = await _tools.pptx_insert_table(path, 1, headers, rows);
        var insertResult = JsonSerializer.Deserialize<TableInsertResult>(result);

        Assert.NotNull(insertResult);
        Assert.True(insertResult.Success);
        Assert.Equal(1, insertResult.SlideNumber);
        Assert.Equal(3, insertResult.RowCount);   // 1 header + 2 data
        Assert.Equal(2, insertResult.ColumnCount);
    }

    [Fact]
    public async Task pptx_insert_table_ReturnsTableName_WhenSpecified()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition { TitleText = "Named" });
        var headers = new[] { "Col1" };
        var rows = new[] { new[] { "Val1" } };

        var result = await _tools.pptx_insert_table(path, 1, headers, rows, tableName: "My Table");
        var insertResult = JsonSerializer.Deserialize<TableInsertResult>(result);

        Assert.NotNull(insertResult);
        Assert.True(insertResult.Success);
        Assert.Equal("My Table", insertResult.TableName);
    }

    // ────────────────────────────────────────────────────────
    // File-not-found: both table tools
    // ────────────────────────────────────────────────────────

    [Theory]
    [InlineData("pptx_insert_table")]
    [InlineData("pptx_update_table")]
    public async Task FileNotFound_ReturnsError(string toolName)
    {
        var fakePath = "C:\\does-not-exist\\file.pptx";
        var result = toolName switch
        {
            "pptx_insert_table" => await _tools.pptx_insert_table(fakePath, 1, ["A"], [["1"]]),
            "pptx_update_table" => await _tools.pptx_update_table(fakePath, 1,
                tableName: "Missing", updates: [new TableCellUpdate(0, 0, "X")]),
            _ => throw new ArgumentException($"Unknown tool: {toolName}")
        };

        var isErrorString = result.StartsWith("Error:", StringComparison.OrdinalIgnoreCase);
        if (isErrorString)
        {
            Assert.Contains("not found", result, StringComparison.OrdinalIgnoreCase);
        }
        else
        {
            using var doc = JsonDocument.Parse(result);
            Assert.False(doc.RootElement.GetProperty("Success").GetBoolean());
            Assert.Contains("not found",
                doc.RootElement.GetProperty("Message").GetString()!,
                StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public async Task pptx_insert_table_InvalidSlideNumber_ReturnsError()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition { TitleText = "Single" });
        var headers = new[] { "A" };
        var rows = new[] { new[] { "1" } };

        var result = await _tools.pptx_insert_table(path, 99, headers, rows);

        // Should indicate failure — either structured JSON or error string
        var isSuccess = false;
        try
        {
            var insertResult = JsonSerializer.Deserialize<TableInsertResult>(result);
            isSuccess = insertResult?.Success ?? false;
        }
        catch (JsonException)
        {
            // Error string format
        }
        Assert.False(isSuccess);
    }

    // ────────────────────────────────────────────────────────
    // pptx_update_table: JSON output
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task pptx_update_table_ReturnsStructuredJson()
    {
        var path = CreatePptxWithTable("Revenue", [["Region", "ARR"], ["NA", "3.2M"]]);

        var result = await _tools.pptx_update_table(path, 1,
            tableName: "Revenue",
            updates: [new TableCellUpdate(1, 1, "4.8M")]);
        var updateResult = JsonSerializer.Deserialize<TableUpdateResult>(result);

        Assert.NotNull(updateResult);
        Assert.True(updateResult.Success);
        Assert.Equal(1, updateResult.SlideNumber);
    }

    [Fact]
    public async Task pptx_update_table_TableNotFound_ReturnsFailure()
    {
        var path = CreatePptxWithTable("Actual", [["A"], ["1"]]);

        var result = await _tools.pptx_update_table(path, 1,
            tableName: "Missing Table",
            updates: [new TableCellUpdate(0, 0, "X")]);

        var isSuccess = false;
        try
        {
            var updateResult = JsonSerializer.Deserialize<TableUpdateResult>(result);
            isSuccess = updateResult?.Success ?? false;
        }
        catch (JsonException)
        {
            // Error string format — still counts as failure
        }
        Assert.False(isSuccess);
    }

    [Fact]
    public async Task pptx_update_table_VerifiesCellContent_AfterUpdate()
    {
        var path = CreatePptxWithTable("KPI Table",
            [["Metric", "Value"], ["ARR", "3.2M"], ["NRR", "112%"]]);

        await _tools.pptx_update_table(path, 1,
            tableName: "KPI Table",
            updates:
            [
                new TableCellUpdate(1, 1, "4.8M"),
                new TableCellUpdate(2, 1, "118%")
            ]);

        // Verify via service-level read
        var slideContent = Service.GetSlideContent(path, 0);
        var tableShape = Assert.Single(slideContent.Shapes, s => s.ShapeType == "Table");
        Assert.Equal("4.8M", tableShape.TableRows![1][1]);
        Assert.Equal("118%", tableShape.TableRows[2][1]);
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
