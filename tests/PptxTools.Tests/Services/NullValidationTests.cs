using System.Text.Json;

namespace PptxTools.Tests.Services;

/// <summary>
/// Verifies that MCP tools handle null and empty inputs gracefully —
/// returning descriptive error strings or structured failure results,
/// never throwing unhandled exceptions to the caller.
/// </summary>
public class NullValidationTests : PptxTestBase
{
    private readonly global::PptxTools.Tools.PptxTools _tools;

    public NullValidationTests()
    {
        _tools = new global::PptxTools.Tools.PptxTools(Service);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Null / empty filePath → every tool should return an error string
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task ListSlides_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_list_slides(filePath!);
        Assert.Contains("Error", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task GetSlideContent_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_get_slide_content(filePath!, 0);
        Assert.Contains("Error", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task GetSlideXml_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_get_slide_xml(filePath!, 0);
        Assert.Contains("Error", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task ExtractTalkingPoints_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_extract_talking_points(filePath!);
        Assert.Contains("Error", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task ExportMarkdown_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_export_markdown(filePath!);
        Assert.Contains("Error", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task UpdateSlideData_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_update_slide_data(filePath!, 1, "Shape", null, "text");
        var parsed = JsonSerializer.Deserialize<SlideDataUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task BatchUpdate_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var mutations = new[] { new BatchUpdateMutation(1, "Title", "val") };
        var result = await _tools.pptx_batch_update(filePath!, mutations);
        Assert.Contains("Error", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task InsertTable_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_insert_table(filePath!, 1, ["A"], [["1"]]);
        var parsed = JsonSerializer.Deserialize<TableInsertResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task UpdateTable_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var updates = new[] { new TableCellUpdate(0, 0, "val") };
        var result = await _tools.pptx_update_table(filePath!, 1, updates);
        var parsed = JsonSerializer.Deserialize<TableUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task InsertImage_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_insert_image(filePath!, 0, "fake.png");
        Assert.Contains("Error", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task ReplaceImage_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_replace_image(filePath!, 1, "Shape", null, "fake.png");
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task WriteNotes_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_write_notes(filePath!, 0, "notes");
        Assert.Contains("Error", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task MoveSlide_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_reorder_slides(filePath!, ReorderSlidesAction.Move, slideNumber: 1, targetPosition: 2);
        Assert.Contains("not found", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task DeleteSlide_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_delete_slide(filePath!, 1);
        Assert.Contains("Error", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task ReorderSlides_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_reorder_slides(filePath!, ReorderSlidesAction.Reorder, newOrder: [1]);
        Assert.Contains("not found", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task ListLayouts_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_list_layouts(filePath!);
        Assert.Contains("Error", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task AddSlide_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_manage_slides(filePath!, ManageSlidesAction.Add);
        Assert.Contains("not found", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task ManageLayouts_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_manage_layouts(filePath!, ManageLayoutsAction.Find);
        Assert.Contains("not found", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task ManageMedia_NullOrEmptyFilePath_ReturnsError(string? filePath)
    {
        var result = await _tools.pptx_manage_media(filePath!, ManageMediaAction.Analyze);
        Assert.Contains("not found", result, StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Null shapeName with null shapeIndex → structured failure
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task UpdateSlideData_NullShapeNameAndNullIndex_ReturnsFailure()
    {
        var path = CreateMinimalPptx();
        var result = await _tools.pptx_update_slide_data(path, 1, null, null, "text");
        var parsed = JsonSerializer.Deserialize<SlideDataUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("shapeName", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UpdateSlideData_Service_NullShapeNameAndNullIndex_ReturnsFailure()
    {
        var path = CreateMinimalPptx();
        var result = Service.UpdateSlideData(path, 1, null, null, "new text");
        Assert.False(result.Success);
        Assert.Contains("shapeName", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Empty mutations array for batch update
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task BatchUpdate_NullOrEmptyMutations_ReturnsEmptyResult(bool useNull)
    {
        var path = CreateMinimalPptx();
        var mutations = useNull ? null! : Array.Empty<BatchUpdateMutation>();
        var result = await _tools.pptx_batch_update(path, mutations);
        var parsed = JsonSerializer.Deserialize<BatchUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.Equal(0, parsed.TotalMutations);
        Assert.Equal(0, parsed.SuccessCount);
        Assert.Equal(0, parsed.FailureCount);
        Assert.Empty(parsed.Results);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void BatchUpdate_Service_NullOrEmptyMutations_ReturnsEmptyResult(bool useNull)
    {
        var path = CreateMinimalPptx();
        var mutations = useNull ? null! : Array.Empty<BatchUpdateMutation>();
        var result = Service.BatchUpdate(path, mutations);
        Assert.Equal(0, result.TotalMutations);
        Assert.Equal(0, result.SuccessCount);
        Assert.Empty(result.Results);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Null headers / rows for table insert
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task InsertTable_NullOrEmptyHeaders_ReturnsFailure(bool useNull)
    {
        var path = CreateMinimalPptx();
        var headers = useNull ? null! : Array.Empty<string>();
        var result = await _tools.pptx_insert_table(path, 1, headers, [["1"]]);
        var parsed = JsonSerializer.Deserialize<TableInsertResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task InsertTable_NullRows_Succeeds()
    {
        // Null rows coalesced to empty — header-only table is valid
        var path = CreateMinimalPptx();
        var result = await _tools.pptx_insert_table(path, 1, ["Col1"], null!);
        var parsed = JsonSerializer.Deserialize<TableInsertResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal(1, parsed.RowCount); // header row only
    }

    // ────────────────────────────────────────────────────────────────────────
    // Null / empty updates for table update
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task UpdateTable_NullOrEmptyUpdates_ReturnsFailure(bool useNull)
    {
        var path = CreateMinimalPptx();
        var updates = useNull ? null! : Array.Empty<TableCellUpdate>();
        var result = await _tools.pptx_update_table(path, 1, updates);
        var parsed = JsonSerializer.Deserialize<TableUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("No updates provided", parsed.Message);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Null imagePath for image operations
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task InsertImage_NullOrEmptyImagePath_ReturnsError(string? imagePath)
    {
        var path = CreateMinimalPptx();
        var result = await _tools.pptx_insert_image(path, 0, imagePath!);
        Assert.Contains("Error", result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task ReplaceImage_NullOrEmptyImagePath_ReturnsFailure(string? imagePath)
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Slide",
            IncludeImage = true
        });
        var result = await _tools.pptx_replace_image(path, 1, null, 0, imagePath!);
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Null / empty shapeName with null shapeIndex for replace image
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task ReplaceImage_NullShapeNameAndNullIndex_ReturnsFailure()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Slide",
            IncludeImage = true
        });
        // Create a real image file so we get past file-exists checks
        var imgPath = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".png");
        TrackTempFile(imgPath);
        File.WriteAllBytes(imgPath, CreateMinimalPng());

        var result = await _tools.pptx_replace_image(path, 1, null, null, imgPath);
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Empty string for text updates
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task UpdateSlideData_EmptyNewText_Succeeds()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Original Title",
            TextShapes =
            [
                new TestTextShapeDefinition
                {
                    Name = "Content",
                    Paragraphs = ["old"]
                }
            ]
        });
        var result = await _tools.pptx_update_slide_data(path, 1, "Content", null, "");
        var parsed = JsonSerializer.Deserialize<SlideDataUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal("", parsed.NewText);
    }

    [Fact]
    public void UpdateSlideData_Service_EmptyNewText_ClearsShape()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Title",
            TextShapes =
            [
                new TestTextShapeDefinition
                {
                    Name = "Content",
                    Paragraphs = ["some text"]
                }
            ]
        });
        var result = Service.UpdateSlideData(path, 1, "Content", null, "");
        Assert.True(result.Success);

        // Verify shape text is now empty
        var content = Service.GetSlideContent(path, 0);
        var shape = content.Shapes.First(s => s.Name == "Content");
        Assert.Equal("", shape.Text);
    }

    [Fact]
    public async Task BatchUpdate_EmptyNewValue_Succeeds()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Title",
            TextShapes =
            [
                new TestTextShapeDefinition
                {
                    Name = "Content",
                    Paragraphs = ["existing"]
                }
            ]
        });
        var mutations = new[] { new BatchUpdateMutation(1, "Content", "") };
        var result = await _tools.pptx_batch_update(path, mutations);
        var parsed = JsonSerializer.Deserialize<BatchUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.Equal(1, parsed.SuccessCount);
        Assert.Equal(0, parsed.FailureCount);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Non-existent file paths (distinct from null/empty — file just missing)
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task ListSlides_NonexistentFile_ReturnsError()
    {
        var result = await _tools.pptx_list_slides(@"C:\no\such\file.pptx");
        Assert.Contains("Error", result);
        Assert.Contains("not found", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task BatchUpdate_NonexistentFile_ReturnsStructuredFailure()
    {
        var mutations = new[] { new BatchUpdateMutation(1, "Title", "val") };
        var result = await _tools.pptx_batch_update(@"C:\no\such\file.pptx", mutations);
        Assert.Contains("not found", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task InsertTable_NonexistentFile_ReturnsStructuredFailure()
    {
        var result = await _tools.pptx_insert_table(@"C:\no\such\file.pptx", 1, ["A"], [["1"]]);
        var parsed = JsonSerializer.Deserialize<TableInsertResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("not found", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ReplaceImage_NonexistentFile_ReturnsStructuredFailure()
    {
        var result = await _tools.pptx_replace_image(@"C:\no\such\file.pptx", 1, "Shape", null, "fake.png");
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("not found", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Helpers
    // ────────────────────────────────────────────────────────────────────────

    /// <summary>Produces a minimal valid 1×1 PNG (67 bytes).</summary>
    private static byte[] CreateMinimalPng()
    {
        using var ms = new MemoryStream();
        // PNG signature
        ms.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
        // IHDR chunk (1×1, 8-bit RGB)
        WriteChunk(ms, "IHDR"u8, [0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0]);
        // IDAT chunk (minimal compressed data for 1 pixel)
        WriteChunk(ms, "IDAT"u8, [0x08, 0xD7, 0x63, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x04, 0x00, 0x01]);
        // IEND chunk
        WriteChunk(ms, "IEND"u8, []);
        return ms.ToArray();
    }

    private static void WriteChunk(Stream s, ReadOnlySpan<byte> type, byte[] data)
    {
        var length = BitConverter.GetBytes(data.Length);
        if (BitConverter.IsLittleEndian) Array.Reverse(length);
        s.Write(length);
        s.Write(type);
        s.Write(data);
        // CRC placeholder (PNG decoders tolerate incorrect CRCs for our test purposes)
        s.Write(new byte[4]);
    }
}
