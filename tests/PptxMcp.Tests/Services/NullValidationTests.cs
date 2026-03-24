using System.Text.Json;

namespace PptxMcp.Tests.Services;

/// <summary>
/// Verifies that MCP tools handle null and empty inputs gracefully —
/// returning descriptive error strings or structured failure results,
/// never throwing unhandled exceptions to the caller.
/// </summary>
public class NullValidationTests : IDisposable
{
    private readonly PresentationService _service = new();
    private readonly PptxTools _tools;
    private readonly List<string> _tempFiles = [];

    public NullValidationTests()
    {
        _tools = new PptxTools(_service);
    }

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
    // Null / empty filePath → every tool should return an error string
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task ListSlides_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_list_slides(null!);
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task ListSlides_EmptyFilePath_ReturnsError()
    {
        var result = await _tools.pptx_list_slides("");
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task GetSlideContent_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_get_slide_content(null!, 0);
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task GetSlideContent_EmptyFilePath_ReturnsError()
    {
        var result = await _tools.pptx_get_slide_content("", 0);
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task GetSlideXml_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_get_slide_xml(null!, 0);
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task GetSlideXml_EmptyFilePath_ReturnsError()
    {
        var result = await _tools.pptx_get_slide_xml("", 0);
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task ExtractTalkingPoints_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_extract_talking_points(null!);
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task ExtractTalkingPoints_EmptyFilePath_ReturnsError()
    {
        var result = await _tools.pptx_extract_talking_points("");
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task ExportMarkdown_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_export_markdown(null!);
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task ExportMarkdown_EmptyFilePath_ReturnsError()
    {
        var result = await _tools.pptx_export_markdown("");
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task UpdateSlideData_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_update_slide_data(null!, 1, "Shape", null, "text");
        var parsed = JsonSerializer.Deserialize<SlideDataUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task UpdateSlideData_EmptyFilePath_ReturnsError()
    {
        var result = await _tools.pptx_update_slide_data("", 1, "Shape", null, "text");
        var parsed = JsonSerializer.Deserialize<SlideDataUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task BatchUpdate_NullFilePath_ReturnsError()
    {
        var mutations = new[] { new BatchUpdateMutation(1, "Title", "val") };
        var result = await _tools.pptx_batch_update(null!, mutations);
        Assert.Contains("Error", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task BatchUpdate_EmptyFilePath_ReturnsError()
    {
        var mutations = new[] { new BatchUpdateMutation(1, "Title", "val") };
        var result = await _tools.pptx_batch_update("", mutations);
        Assert.Contains("Error", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task InsertTable_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_insert_table(null!, 1, ["A"], [["1"]]);
        var parsed = JsonSerializer.Deserialize<TableInsertResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task InsertTable_EmptyFilePath_ReturnsError()
    {
        var result = await _tools.pptx_insert_table("", 1, ["A"], [["1"]]);
        var parsed = JsonSerializer.Deserialize<TableInsertResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task UpdateTable_NullFilePath_ReturnsError()
    {
        var updates = new[] { new TableCellUpdate(0, 0, "val") };
        var result = await _tools.pptx_update_table(null!, 1, updates);
        var parsed = JsonSerializer.Deserialize<TableUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task UpdateTable_EmptyFilePath_ReturnsError()
    {
        var updates = new[] { new TableCellUpdate(0, 0, "val") };
        var result = await _tools.pptx_update_table("", 1, updates);
        var parsed = JsonSerializer.Deserialize<TableUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task InsertImage_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_insert_image(null!, 0, "fake.png");
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task InsertImage_EmptyFilePath_ReturnsError()
    {
        var result = await _tools.pptx_insert_image("", 0, "fake.png");
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task ReplaceImage_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_replace_image(null!, 1, "Shape", null, "fake.png");
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task ReplaceImage_EmptyFilePath_ReturnsError()
    {
        var result = await _tools.pptx_replace_image("", 1, "Shape", null, "fake.png");
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task WriteNotes_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_write_notes(null!, 0, "notes");
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task WriteNotes_EmptyFilePath_ReturnsError()
    {
        var result = await _tools.pptx_write_notes("", 0, "notes");
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task MoveSlide_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_reorder_slides(null!, ReorderSlidesAction.Move, slideNumber: 1, targetPosition: 2);
        Assert.Contains("File not found", result);
    }

    [Fact]
    public async Task DeleteSlide_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_delete_slide(null!, 1);
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task ReorderSlides_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_reorder_slides(null!, ReorderSlidesAction.Reorder, newOrder: [1]);
        Assert.Contains("File not found", result);
    }

    [Fact]
    public async Task ListLayouts_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_list_layouts(null!);
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task AddSlide_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_manage_slides(null!, ManageSlidesAction.Add);
        Assert.Contains("File not found", result);
    }

    [Fact]
    public async Task ManageLayouts_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_manage_layouts(null!, ManageLayoutsAction.Find);
        Assert.Contains("File not found", result);
    }

    [Fact]
    public async Task ManageMedia_NullFilePath_ReturnsError()
    {
        var result = await _tools.pptx_manage_media(null!, ManageMediaAction.Analyze);
        Assert.Contains("File not found", result);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Null shapeName with null shapeIndex → structured failure
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task UpdateSlideData_NullShapeNameAndNullIndex_ReturnsFailure()
    {
        var path = CreateMinimalPresentation();
        var result = await _tools.pptx_update_slide_data(path, 1, null, null, "text");
        var parsed = JsonSerializer.Deserialize<SlideDataUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("shapeName", parsed.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UpdateSlideData_Service_NullShapeNameAndNullIndex_ReturnsFailure()
    {
        var path = CreateMinimalPresentation();
        var result = _service.UpdateSlideData(path, 1, null, null, "new text");
        Assert.False(result.Success);
        Assert.Contains("shapeName", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Empty mutations array for batch update
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task BatchUpdate_NullMutations_ReturnsEmptyResult()
    {
        var path = CreateMinimalPresentation();
        var result = await _tools.pptx_batch_update(path, null!);
        var parsed = JsonSerializer.Deserialize<BatchUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.Equal(0, parsed.TotalMutations);
        Assert.Equal(0, parsed.SuccessCount);
        Assert.Equal(0, parsed.FailureCount);
        Assert.Empty(parsed.Results);
    }

    [Fact]
    public async Task BatchUpdate_EmptyMutations_ReturnsEmptyResult()
    {
        var path = CreateMinimalPresentation();
        var result = await _tools.pptx_batch_update(path, []);
        var parsed = JsonSerializer.Deserialize<BatchUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.Equal(0, parsed.TotalMutations);
        Assert.Equal(0, parsed.SuccessCount);
        Assert.Equal(0, parsed.FailureCount);
        Assert.Empty(parsed.Results);
    }

    [Fact]
    public void BatchUpdate_Service_NullMutations_ReturnsEmptyResult()
    {
        var path = CreateMinimalPresentation();
        var result = _service.BatchUpdate(path, null!);
        Assert.Equal(0, result.TotalMutations);
        Assert.Equal(0, result.SuccessCount);
        Assert.Empty(result.Results);
    }

    [Fact]
    public void BatchUpdate_Service_EmptyMutations_ReturnsEmptyResult()
    {
        var path = CreateMinimalPresentation();
        var result = _service.BatchUpdate(path, []);
        Assert.Equal(0, result.TotalMutations);
        Assert.Equal(0, result.SuccessCount);
        Assert.Empty(result.Results);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Null headers / rows for table insert
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task InsertTable_NullHeaders_ReturnsFailure()
    {
        var path = CreateMinimalPresentation();
        var result = await _tools.pptx_insert_table(path, 1, null!, [["1"]]);
        var parsed = JsonSerializer.Deserialize<TableInsertResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task InsertTable_EmptyHeaders_ReturnsFailure()
    {
        var path = CreateMinimalPresentation();
        var result = await _tools.pptx_insert_table(path, 1, [], [["1"]]);
        var parsed = JsonSerializer.Deserialize<TableInsertResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task InsertTable_NullRows_Succeeds()
    {
        // Null rows coalesced to empty — header-only table is valid
        var path = CreateMinimalPresentation();
        var result = await _tools.pptx_insert_table(path, 1, ["Col1"], null!);
        var parsed = JsonSerializer.Deserialize<TableInsertResult>(result);
        Assert.NotNull(parsed);
        Assert.True(parsed.Success);
        Assert.Equal(1, parsed.RowCount); // header row only
    }

    // ────────────────────────────────────────────────────────────────────────
    // Null / empty updates for table update
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task UpdateTable_NullUpdates_ReturnsFailure()
    {
        var path = CreateMinimalPresentation();
        var result = await _tools.pptx_update_table(path, 1, null!);
        var parsed = JsonSerializer.Deserialize<TableUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("No updates provided", parsed.Message);
    }

    [Fact]
    public async Task UpdateTable_EmptyUpdates_ReturnsFailure()
    {
        var path = CreateMinimalPresentation();
        var result = await _tools.pptx_update_table(path, 1, []);
        var parsed = JsonSerializer.Deserialize<TableUpdateResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
        Assert.Contains("No updates provided", parsed.Message);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Null imagePath for image operations
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public async Task InsertImage_NullImagePath_ReturnsError()
    {
        var path = CreateMinimalPresentation();
        var result = await _tools.pptx_insert_image(path, 0, null!);
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task InsertImage_EmptyImagePath_ReturnsError()
    {
        var path = CreateMinimalPresentation();
        var result = await _tools.pptx_insert_image(path, 0, "");
        Assert.Contains("Error", result);
    }

    [Fact]
    public async Task ReplaceImage_NullImagePath_ReturnsFailure()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Slide",
            IncludeImage = true
        });
        var result = await _tools.pptx_replace_image(path, 1, null, 0, null!);
        var parsed = JsonSerializer.Deserialize<ImageReplaceResult>(result);
        Assert.NotNull(parsed);
        Assert.False(parsed.Success);
    }

    [Fact]
    public async Task ReplaceImage_EmptyImagePath_ReturnsFailure()
    {
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Slide",
            IncludeImage = true
        });
        var result = await _tools.pptx_replace_image(path, 1, null, 0, "");
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
        var path = CreatePresentation(new TestSlideDefinition
        {
            TitleText = "Slide",
            IncludeImage = true
        });
        // Create a real image file so we get past file-exists checks
        var imgPath = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".png");
        _tempFiles.Add(imgPath);
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
        var path = CreatePresentation(new TestSlideDefinition
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
        var path = CreatePresentation(new TestSlideDefinition
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
        var result = _service.UpdateSlideData(path, 1, "Content", null, "");
        Assert.True(result.Success);

        // Verify shape text is now empty
        var content = _service.GetSlideContent(path, 0);
        var shape = content.Shapes.First(s => s.Name == "Content");
        Assert.Equal("", shape.Text);
    }

    [Fact]
    public async Task BatchUpdate_EmptyNewValue_Succeeds()
    {
        var path = CreatePresentation(new TestSlideDefinition
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
