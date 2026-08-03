using System.Text.Json;
using PptxTools.Models;

namespace PptxTools.Tests.Tools;

[Trait("Category", "Integration")]
public class BatchExecuteToolsTests : PptxTestBase
{
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public BatchExecuteToolsTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    // ────────────────────────────────────────────────────────
    // JSON serialization of BatchOperationResult
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task pptx_batch_execute_ReturnsStructuredJson()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "JSON Test",
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "Target", Paragraphs = ["Original"] }
            ]
        });

        var json = await _tools.pptx_batch_execute(path,
        [
            new BatchOperation(1, "Target", BatchOperationType.UpdateText, NewText: "Updated")
        ]);

        var result = JsonSerializer.Deserialize<BatchOperationResult>(json);

        Assert.NotNull(result);
        Assert.Equal(1, result.TotalOperations);
        Assert.Equal(1, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
        Assert.False(result.RolledBack);
        Assert.Single(result.Results);
    }

    [Fact]
    public async Task pptx_batch_execute_UpdateText_JsonHasCorrectStructure()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Slide",
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "Shape1", Paragraphs = ["Hello"] }
            ]
        });

        var json = await _tools.pptx_batch_execute(path,
        [
            new BatchOperation(1, "Shape1", BatchOperationType.UpdateText, NewText: "World")
        ]);

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.Equal(1, root.GetProperty("TotalOperations").GetInt32());
        Assert.Equal(1, root.GetProperty("SuccessCount").GetInt32());
        Assert.Equal(0, root.GetProperty("FailureCount").GetInt32());
        Assert.False(root.GetProperty("RolledBack").GetBoolean());

        var results = root.GetProperty("Results");
        Assert.Equal(1, results.GetArrayLength());
        var first = results[0];
        Assert.Equal(1, first.GetProperty("SlideNumber").GetInt32());
        Assert.Equal("Shape1", first.GetProperty("ShapeName").GetString());
        Assert.True(first.GetProperty("Success").GetBoolean());
    }

    [Fact]
    public async Task pptx_batch_execute_UpdateTableCell_JsonHasCorrectStructure()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Table Test",
            Tables =
            [
                new TestTableDefinition
                {
                    Name = "MyTable",
                    Rows = [["A", "B"], ["1", "2"]]
                }
            ]
        });

        var json = await _tools.pptx_batch_execute(path,
        [
            new BatchOperation(1, "MyTable", BatchOperationType.UpdateTableCell,
                TableRow: 1, TableColumn: 0, CellValue: "Updated")
        ]);

        using var doc = JsonDocument.Parse(json);
        var outcome = doc.RootElement.GetProperty("Results")[0];
        Assert.True(outcome.GetProperty("Success").GetBoolean());
        Assert.Equal((int)BatchOperationType.UpdateTableCell, outcome.GetProperty("Type").GetInt32());
    }

    [Fact]
    public async Task pptx_batch_execute_UpdateShapeProperties_JsonHasCorrectStructure()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Shape Props Test",
            TextShapes =
            [
                new TestTextShapeDefinition
                {
                    Name = "Moveable",
                    Paragraphs = ["Content"],
                    X = Emu.OneInch,
                    Y = Emu.OneInch,
                    Width = Emu.Inches3,
                    Height = Emu.ThreeQuartersInch
                }
            ]
        });

        var json = await _tools.pptx_batch_execute(path,
        [
            new BatchOperation(1, "Moveable", BatchOperationType.UpdateShapeProperties,
                X: Emu.Inches4)
        ]);

        using var doc = JsonDocument.Parse(json);
        var outcome = doc.RootElement.GetProperty("Results")[0];
        Assert.True(outcome.GetProperty("Success").GetBoolean());
        Assert.Equal((int)BatchOperationType.UpdateShapeProperties, outcome.GetProperty("Type").GetInt32());
    }

    [Fact]
    public async Task pptx_batch_execute_ReplaceImage_JsonHasCorrectStructure()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Image Test",
            IncludeImage = true
        });
        var imagePath = CreateTempPng();

        // Use the default picture name from TestPptxHelper
        var json = await _tools.pptx_batch_execute(path,
        [
            new BatchOperation(1, "Picture 3", BatchOperationType.ReplaceImage, ImagePath: imagePath)
        ]);

        using var doc = JsonDocument.Parse(json);
        var outcome = doc.RootElement.GetProperty("Results")[0];
        Assert.True(outcome.GetProperty("Success").GetBoolean());
        Assert.Equal((int)BatchOperationType.ReplaceImage, outcome.GetProperty("Type").GetInt32());
    }

    // ────────────────────────────────────────────────────────
    // File not found
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task pptx_batch_execute_FileNotFound_ReturnsError()
    {
        var json = await _tools.pptx_batch_execute(@"C:\does-not-exist\file.pptx",
        [
            new BatchOperation(1, "Shape", BatchOperationType.UpdateText, NewText: "X")
        ]);

        // Tool returns either "Error: ..." string or structured JSON with all failures
        var isErrorString = json.StartsWith("Error:", StringComparison.OrdinalIgnoreCase);
        if (isErrorString)
        {
            Assert.Contains("not found", json, StringComparison.OrdinalIgnoreCase);
        }
        else
        {
            using var doc = JsonDocument.Parse(json);
            Assert.Equal(0, doc.RootElement.GetProperty("SuccessCount").GetInt32());
        }
    }

    // ────────────────────────────────────────────────────────
    // Empty operations
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task pptx_batch_execute_EmptyOperations_ReturnsZeroCounts()
    {
        // Empty operations should return structured result even without a real file
        var json = await _tools.pptx_batch_execute("anything.pptx", []);

        var result = JsonSerializer.Deserialize<BatchOperationResult>(json);

        Assert.NotNull(result);
        Assert.Equal(0, result.TotalOperations);
        Assert.Equal(0, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
        Assert.False(result.RolledBack);
        Assert.Empty(result.Results);
    }

    [Fact]
    public async Task pptx_batch_execute_NullOperations_ReturnsZeroCounts()
    {
        var json = await _tools.pptx_batch_execute("anything.pptx", null!);

        var result = JsonSerializer.Deserialize<BatchOperationResult>(json);

        Assert.NotNull(result);
        Assert.Equal(0, result.TotalOperations);
    }

    // ────────────────────────────────────────────────────────
    // Atomic flag in JSON output
    // ────────────────────────────────────────────────────────

    [Fact]
    public async Task pptx_batch_execute_AtomicSuccess_RolledBackFalse()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Atomic",
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "S1", Paragraphs = ["V1"] }
            ]
        });

        var json = await _tools.pptx_batch_execute(path,
        [
            new BatchOperation(1, "S1", BatchOperationType.UpdateText, NewText: "V2")
        ], atomic: true);

        using var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.GetProperty("RolledBack").GetBoolean());
        Assert.Equal(1, doc.RootElement.GetProperty("SuccessCount").GetInt32());
    }

    [Fact]
    public async Task pptx_batch_execute_AtomicFailure_RolledBackTrue()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Atomic Fail",
            TextShapes =
            [
                new TestTextShapeDefinition { Name = "S1", Paragraphs = ["V1"] }
            ]
        });

        var json = await _tools.pptx_batch_execute(path,
        [
            new BatchOperation(1, "S1", BatchOperationType.UpdateText, NewText: "Good"),
            new BatchOperation(1, "Missing", BatchOperationType.UpdateText, NewText: "Bad")
        ], atomic: true);

        using var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("RolledBack").GetBoolean());
    }

    // ────────────────────────────────────────────────────────
    // Helpers
    // ────────────────────────────────────────────────────────

    private static readonly byte[] MinimalPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nZxQAAAAASUVORK5CYII=");

    private string CreateTempPng()
    {
        var imagePath = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".png");
        File.WriteAllBytes(imagePath, MinimalPng);
        TrackTempFile(imagePath);
        return imagePath;
    }
}
