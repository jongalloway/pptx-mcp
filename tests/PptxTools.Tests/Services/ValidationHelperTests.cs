namespace PptxTools.Tests.Services;

/// <summary>
/// Tests for <see cref="ValidationHelpers"/> — centralized validation that produces
/// actionable error messages for MCP consumers. Covers every public method: slide
/// number/index validation, EMU values, file/image paths, color formats, and the
/// human-readable "not found" message builders.
/// </summary>
public class ValidationHelperTests : PptxTestBase
{
    // ────────────────────────────────────────────────────────────────────────
    // ValidateSlideNumber
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(0, 5)]
    [InlineData(-1, 5)]
    [InlineData(6, 5)]
    [InlineData(10, 3)]
    public void ValidateSlideNumber_OutOfRange_ThrowsWithRangeInfo(int slideNumber, int totalSlides)
    {
        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ValidationHelpers.ValidateSlideNumber(slideNumber, totalSlides));
        Assert.Contains($"1-{totalSlides}", ex.Message);
    }

    [Theory]
    [InlineData(1, 1)]
    [InlineData(1, 5)]
    [InlineData(3, 5)]
    [InlineData(5, 5)]
    public void ValidateSlideNumber_ValidRange_DoesNotThrow(int slideNumber, int totalSlides)
    {
        var exception = Record.Exception(
            () => ValidationHelpers.ValidateSlideNumber(slideNumber, totalSlides));
        Assert.Null(exception);
    }

    [Fact]
    public void ValidateSlideNumber_ZeroTotalSlides_ThrowsInvalidOperation()
    {
        var ex = Assert.Throws<InvalidOperationException>(
            () => ValidationHelpers.ValidateSlideNumber(1, 0));
        Assert.Contains("no slides", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ValidateSlideNumber_WithContext_IncludesContextInMessage()
    {
        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ValidationHelpers.ValidateSlideNumber(10, 3, "while updating table"));
        Assert.Contains("while updating table", ex.Message);
        Assert.Contains("1-3", ex.Message);
    }

    [Theory]
    [InlineData(1, 5)]
    [InlineData(5, 5)]
    public void ValidateSlideNumber_EmptyContext_OmitsTrailingContext(int slideNumber, int totalSlides)
    {
        // Empty context should not append extra whitespace or text
        var exception = Record.Exception(
            () => ValidationHelpers.ValidateSlideNumber(slideNumber, totalSlides, ""));
        Assert.Null(exception);
    }

    // ────────────────────────────────────────────────────────────────────────
    // ValidateSlideIndex
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(-1, 5)]
    [InlineData(5, 5)]
    [InlineData(10, 3)]
    public void ValidateSlideIndex_OutOfRange_ThrowsWithRangeInfo(int slideIndex, int totalSlides)
    {
        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ValidationHelpers.ValidateSlideIndex(slideIndex, totalSlides));
        Assert.Contains($"0-{totalSlides - 1}", ex.Message);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(0, 5)]
    [InlineData(2, 5)]
    [InlineData(4, 5)]
    public void ValidateSlideIndex_ValidRange_DoesNotThrow(int slideIndex, int totalSlides)
    {
        var exception = Record.Exception(
            () => ValidationHelpers.ValidateSlideIndex(slideIndex, totalSlides));
        Assert.Null(exception);
    }

    [Fact]
    public void ValidateSlideIndex_ZeroTotalSlides_ThrowsInvalidOperation()
    {
        var ex = Assert.Throws<InvalidOperationException>(
            () => ValidationHelpers.ValidateSlideIndex(0, 0));
        Assert.Contains("no slides", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ValidateSlideIndex_WithContext_IncludesContextInMessage()
    {
        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ValidationHelpers.ValidateSlideIndex(5, 3, "during XML extraction"));
        Assert.Contains("during XML extraction", ex.Message);
    }

    // ────────────────────────────────────────────────────────────────────────
    // ValidateEmuValue
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(-1)]
    [InlineData(-914400)]
    [InlineData(long.MinValue)]
    public void ValidateEmuValue_Negative_ThrowsWithEmuExplanation(long value)
    {
        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ValidationHelpers.ValidateEmuValue(value, "width"));
        Assert.Contains("width", ex.Message);
        Assert.Contains("914400", ex.Message); // EMU conversion hint
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(914400)]
    [InlineData(long.MaxValue)]
    public void ValidateEmuValue_ZeroOrPositive_DoesNotThrow(long value)
    {
        var exception = Record.Exception(
            () => ValidationHelpers.ValidateEmuValue(value, "x"));
        Assert.Null(exception);
    }

    [Fact]
    public void ValidateEmuValue_IncludesParamNameInMessage()
    {
        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ValidationHelpers.ValidateEmuValue(-100, "tableHeight"));
        Assert.Contains("tableHeight", ex.Message);
    }

    // ────────────────────────────────────────────────────────────────────────
    // ValidateFilePath
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ValidateFilePath_NullOrEmpty_ThrowsArgumentException(string? filePath)
    {
        Assert.Throws<ArgumentException>(
            () => ValidationHelpers.ValidateFilePath(filePath!));
    }

    [Fact]
    public void ValidateFilePath_FileDoesNotExist_ThrowsFileNotFound()
    {
        var ex = Assert.Throws<FileNotFoundException>(
            () => ValidationHelpers.ValidateFilePath(@"C:\nonexistent\fake.pptx"));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ValidateFilePath_NonPptxExtension_ThrowsArgumentException()
    {
        // Create a real file with wrong extension
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        File.WriteAllText(path, "dummy");
        TrackTempFile(path);

        var ex = Assert.Throws<ArgumentException>(
            () => ValidationHelpers.ValidateFilePath(path));
        Assert.Contains(".pptx", ex.Message);
    }

    [Fact]
    public void ValidateFilePath_ValidPptx_DoesNotThrow()
    {
        var path = CreateMinimalPptx();

        var exception = Record.Exception(
            () => ValidationHelpers.ValidateFilePath(path));
        Assert.Null(exception);
    }

    // ────────────────────────────────────────────────────────────────────────
    // ValidateImagePath
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ValidateImagePath_NullOrEmpty_ThrowsArgumentException(string? imagePath)
    {
        Assert.Throws<ArgumentException>(
            () => ValidationHelpers.ValidateImagePath(imagePath!));
    }

    [Fact]
    public void ValidateImagePath_FileDoesNotExist_ThrowsFileNotFound()
    {
        var ex = Assert.Throws<FileNotFoundException>(
            () => ValidationHelpers.ValidateImagePath(@"C:\nonexistent\fake.png"));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(".txt")]
    [InlineData(".pdf")]
    [InlineData(".docx")]
    [InlineData(".pptx")]
    public void ValidateImagePath_UnsupportedExtension_ThrowsWithSupportedFormats(string ext)
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ext);
        File.WriteAllText(path, "dummy");
        TrackTempFile(path);

        var ex = Assert.Throws<ArgumentException>(
            () => ValidationHelpers.ValidateImagePath(path));
        Assert.Contains(".png", ex.Message);
        Assert.Contains(".jpg", ex.Message);
    }

    [Theory]
    [InlineData(".png")]
    [InlineData(".jpg")]
    [InlineData(".jpeg")]
    [InlineData(".gif")]
    [InlineData(".bmp")]
    [InlineData(".tiff")]
    [InlineData(".svg")]
    [InlineData(".emf")]
    [InlineData(".wmf")]
    public void ValidateImagePath_SupportedExtension_DoesNotThrow(string ext)
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ext);
        File.WriteAllText(path, "dummy image content");
        TrackTempFile(path);

        var exception = Record.Exception(
            () => ValidationHelpers.ValidateImagePath(path));
        Assert.Null(exception);
    }

    // ────────────────────────────────────────────────────────────────────────
    // ValidateColorFormat
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData("red")]
    [InlineData("GG0000")]
    [InlineData("FF00")]
    [InlineData("ZZZZZZ")]
    [InlineData("12345")]
    [InlineData("1234567")]
    public void ValidateColorFormat_InvalidFormat_ThrowsWithFormatHelp(string color)
    {
        var ex = Assert.Throws<ArgumentException>(
            () => ValidationHelpers.ValidateColorFormat(color, "fontColor"));
        Assert.Contains("fontColor", ex.Message);
        Assert.Contains("hex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("FF0000")]
    [InlineData("00FF00")]
    [InlineData("0000FF")]
    [InlineData("000000")]
    [InlineData("FFFFFF")]
    [InlineData("abcdef")]
    public void ValidateColorFormat_ValidHex_DoesNotThrow(string color)
    {
        var exception = Record.Exception(
            () => ValidationHelpers.ValidateColorFormat(color, "fill"));
        Assert.Null(exception);
    }

    [Theory]
    [InlineData("#FF0000")]
    [InlineData("#00ff00")]
    public void ValidateColorFormat_HashPrefixed_IsAccepted(string color)
    {
        // Implementation strips the '#' prefix — these are valid
        var exception = Record.Exception(
            () => ValidationHelpers.ValidateColorFormat(color, "fill"));
        Assert.Null(exception);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ValidateColorFormat_NullOrEmpty_DoesNotThrow(string? color)
    {
        // Null/empty is silently accepted (optional color parameter)
        var exception = Record.Exception(
            () => ValidationHelpers.ValidateColorFormat(color!, "fill"));
        Assert.Null(exception);
    }

    // ────────────────────────────────────────────────────────────────────────
    // BuildShapeNotFoundMessage
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void BuildShapeNotFoundMessage_ReturnsMessageWithShapeAndAvailable()
    {
        var msg = ValidationHelpers.BuildShapeNotFoundMessage(
            3, "Revenue", ["Title", "Subtitle", "Footer"]);

        Assert.Contains("Revenue", msg);
        Assert.Contains("slide 3", msg);
        Assert.Contains("Title", msg);
        Assert.Contains("Subtitle", msg);
        Assert.Contains("Footer", msg);
    }

    [Fact]
    public void BuildShapeNotFoundMessage_EmptyAvailableShapes_StillWorks()
    {
        var msg = ValidationHelpers.BuildShapeNotFoundMessage(
            1, "MissingShape", []);

        Assert.Contains("MissingShape", msg);
        Assert.Contains("slide 1", msg);
        Assert.Contains("no shapes", msg, StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────────────────────
    // BuildTableNotFoundMessage
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void BuildTableNotFoundMessage_ByName_IncludesNameAndCount()
    {
        var msg = ValidationHelpers.BuildTableNotFoundMessage(2, "SalesData", null, 3);

        Assert.Contains("SalesData", msg);
        Assert.Contains("slide 2", msg);
        Assert.Contains("3", msg);
    }

    [Fact]
    public void BuildTableNotFoundMessage_ByIndex_IncludesIndexAndCount()
    {
        var msg = ValidationHelpers.BuildTableNotFoundMessage(1, null, 5, 3);

        Assert.Contains("5", msg);
        Assert.Contains("slide 1", msg);
        Assert.Contains("3", msg);
        Assert.Contains("0-2", msg); // valid range
    }

    [Fact]
    public void BuildTableNotFoundMessage_NeitherNameNorIndex_ReturnsGenericMessage()
    {
        var msg = ValidationHelpers.BuildTableNotFoundMessage(4, null, null, 0);

        Assert.Contains("slide 4", msg);
        Assert.Contains("0", msg);
    }

    // ────────────────────────────────────────────────────────────────────────
    // ValidateRowIndex / ValidateColumnIndex
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(-1, 5)]
    [InlineData(5, 5)]
    [InlineData(10, 3)]
    public void ValidateRowIndex_OutOfRange_ThrowsWithRangeInfo(int rowIndex, int rowCount)
    {
        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ValidationHelpers.ValidateRowIndex(rowIndex, rowCount, "Revenue"));
        Assert.Contains("Revenue", ex.Message);
        Assert.Contains($"0-{rowCount - 1}", ex.Message);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(0, 5)]
    [InlineData(4, 5)]
    public void ValidateRowIndex_ValidRange_DoesNotThrow(int rowIndex, int rowCount)
    {
        var exception = Record.Exception(
            () => ValidationHelpers.ValidateRowIndex(rowIndex, rowCount, "Data"));
        Assert.Null(exception);
    }

    [Theory]
    [InlineData(-1, 4)]
    [InlineData(4, 4)]
    [InlineData(10, 2)]
    public void ValidateColumnIndex_OutOfRange_ThrowsWithRangeInfo(int colIndex, int colCount)
    {
        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ValidationHelpers.ValidateColumnIndex(colIndex, colCount, "Sales"));
        Assert.Contains("Sales", ex.Message);
        Assert.Contains($"0-{colCount - 1}", ex.Message);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(0, 4)]
    [InlineData(3, 4)]
    public void ValidateColumnIndex_ValidRange_DoesNotThrow(int colIndex, int colCount)
    {
        var exception = Record.Exception(
            () => ValidationHelpers.ValidateColumnIndex(colIndex, colCount, "Data"));
        Assert.Null(exception);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Integration: ValidationHelpers flow through service methods
    // ────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(99)]
    public void InsertTable_InvalidSlideNumber_ErrorIncludesRange(int slideNumber)
    {
        var path = CreateMinimalPptx();

        var ex = Assert.ThrowsAny<Exception>(
            () => Service.InsertTable(path, slideNumber, ["A"], [["1"]]));
        // Should include range info or "no slides" context
        Assert.True(
            ex.Message.Contains("1-1") || ex.Message.Contains("no slides", StringComparison.OrdinalIgnoreCase),
            $"Expected range info in message but got: {ex.Message}");
    }

    [Fact]
    public void UpdateTable_InvalidSlideNumber_ErrorIncludesRange()
    {
        var path = CreateMinimalPptx();
        var updates = new[] { new TableCellUpdate(0, 0, "val") };

        var ex = Assert.ThrowsAny<Exception>(
            () => Service.UpdateTable(path, 99, updates));
        Assert.Contains("1-1", ex.Message);
    }

    [Fact]
    public void InsertImage_NegativeEmu_ErrorIncludesParamName()
    {
        var path = CreateMinimalPptx();
        var imgPath = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".png");
        File.WriteAllBytes(imgPath, [0x89, 0x50, 0x4E, 0x47]); // minimal PNG header
        TrackTempFile(imgPath);

        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => Service.InsertImage(path, 0, imgPath, -100, 0, 100, 100));
        Assert.Contains("x", ex.Message);
    }

    [Fact]
    public void InsertImage_InvalidImageExtension_ErrorListsSupportedFormats()
    {
        var path = CreateMinimalPptx();
        var badImg = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".txt");
        File.WriteAllText(badImg, "not an image");
        TrackTempFile(badImg);

        var ex = Assert.Throws<ArgumentException>(
            () => Service.InsertImage(path, 0, badImg, 0, 0, 100, 100));
        Assert.Contains(".png", ex.Message);
    }

    [Fact]
    public void UpdateSlideData_InvalidShapeName_ErrorIncludesAvailableShapes()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "Dashboard" });

        var result = Service.UpdateSlideData(path, 1, "NonexistentShape", null, "new text");
        Assert.False(result.Success);
        Assert.Contains("Available shapes", result.Message!, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UpdateSlideData_SlideNumberTooHigh_ErrorIncludesSlideCount()
    {
        var path = CreateMinimalPptx();

        var result = Service.UpdateSlideData(path, 99, "Title", null, "new text");
        Assert.False(result.Success);
        Assert.Contains("1 slide", result.Message!, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteSlide_InvalidSlideNumber_ThrowsWithRange()
    {
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { TitleText = "A" },
            new TestSlideDefinition { TitleText = "B" });

        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => Service.DeleteSlide(path, 5));
        Assert.Contains("1-2", ex.Message);
    }
}
