namespace PptxTools.Services;

/// <summary>
/// Centralized validation helpers that produce actionable error messages for MCP consumers.
/// All methods throw standard exceptions caught by the ExecuteToolStructured pattern in the tool layer.
/// </summary>
public static class ValidationHelpers
{
    /// <summary>Validate a 1-based slide number is in range.</summary>
    public static void ValidateSlideNumber(int slideNumber, int totalSlides, string context = "")
    {
        if (totalSlides == 0)
            throw new InvalidOperationException("Presentation has no slides.");

        if (slideNumber < 1 || slideNumber > totalSlides)
        {
            var msg = $"Slide {slideNumber} does not exist — out of range. Valid range: 1-{totalSlides}.";
            if (!string.IsNullOrWhiteSpace(context))
                msg += $" {context}";
            throw new ArgumentOutOfRangeException(nameof(slideNumber), msg);
        }
    }

    /// <summary>Validate a 0-based slide index is in range.</summary>
    public static void ValidateSlideIndex(int slideIndex, int totalSlides, string context = "")
    {
        if (totalSlides == 0)
            throw new InvalidOperationException("Presentation has no slides.");

        if (slideIndex < 0 || slideIndex >= totalSlides)
        {
            var msg = $"Slide index {slideIndex} is out of range. Presentation has {totalSlides} slide(s), valid range: 0-{totalSlides - 1}.";
            if (!string.IsNullOrWhiteSpace(context))
                msg += $" {context}";
            throw new ArgumentOutOfRangeException(nameof(slideIndex), msg);
        }
    }

    /// <summary>Validate EMU coordinates (position/size) are non-negative.</summary>
    public static void ValidateEmuValue(long value, string paramName)
    {
        if (value < 0)
            throw new ArgumentOutOfRangeException(paramName,
                $"EMU value for '{paramName}' must be non-negative. Got: {value}. (1 inch = 914400 EMU)");
    }

    /// <summary>Validate a .pptx file path exists and has the correct extension.</summary>
    public static void ValidateFilePath(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path must not be empty.", nameof(filePath));

        if (!File.Exists(filePath))
            throw new FileNotFoundException($"File not found: '{filePath}'. Verify the path is correct and the file exists.", filePath);

        var ext = Path.GetExtension(filePath);
        if (!string.Equals(ext, ".pptx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"Expected a .pptx file but got '{ext}'. Provide a PowerPoint (.pptx) file.", nameof(filePath));
    }

    private static readonly HashSet<string> SupportedImageExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".svg", ".emf", ".wmf"
    };

    /// <summary>Validate an image file path exists and has a supported format.</summary>
    public static void ValidateImagePath(string imagePath)
    {
        if (string.IsNullOrWhiteSpace(imagePath))
            throw new ArgumentException("Image path must not be empty.", nameof(imagePath));

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: '{imagePath}'. Verify the path is correct.", imagePath);

        var ext = Path.GetExtension(imagePath);
        if (!SupportedImageExtensions.Contains(ext))
            throw new ArgumentException(
                $"Unsupported image format '{ext}'. Supported: .png, .jpg, .jpeg, .gif, .bmp, .tiff, .svg, .emf, .wmf",
                nameof(imagePath));
    }

    /// <summary>Validate a 6-digit hex color string (with or without '#' prefix).</summary>
    public static void ValidateColorFormat(string color, string paramName)
    {
        if (string.IsNullOrWhiteSpace(color))
            return;

        var hex = color.StartsWith('#') ? color[1..] : color;
        if (!System.Text.RegularExpressions.Regex.IsMatch(hex, @"^[0-9A-Fa-f]{6}$"))
            throw new ArgumentException(
                $"Invalid color format for '{paramName}': '{color}'. Expected 6-digit hex (e.g., 'FF0000' for red, '0000FF' for blue).",
                paramName);
    }

    /// <summary>Build an actionable "shape not found" message listing available shapes.</summary>
    public static string BuildShapeNotFoundMessage(int slideNumber, string requestedShape, IEnumerable<string> availableShapes)
    {
        var shapeList = string.Join(", ", availableShapes);
        return string.IsNullOrEmpty(shapeList)
            ? $"Shape '{requestedShape}' not found on slide {slideNumber}. The slide has no shapes."
            : $"Shape '{requestedShape}' not found on slide {slideNumber}. Available shapes: {shapeList}";
    }

    /// <summary>Build an actionable "table not found" message with count context.</summary>
    public static string BuildTableNotFoundMessage(int slideNumber, string? tableName, int? tableIndex, int tableCount)
    {
        if (tableName is not null)
            return $"Table '{tableName}' not found on slide {slideNumber}. Found {tableCount} table(s).";

        if (tableIndex.HasValue)
            return $"Table index {tableIndex.Value} is out of range on slide {slideNumber}. Found {tableCount} table(s), valid range: 0-{tableCount - 1}.";

        return $"No table found on slide {slideNumber}. Found {tableCount} table(s).";
    }

    /// <summary>Validate a row index is within a table's row count.</summary>
    public static void ValidateRowIndex(int rowIndex, int rowCount, string tableName)
    {
        if (rowIndex < 0 || rowIndex >= rowCount)
            throw new ArgumentOutOfRangeException(nameof(rowIndex),
                $"Row index {rowIndex} is out of range. Table '{tableName}' has {rowCount} row(s), valid range: 0-{rowCount - 1}.");
    }

    /// <summary>Validate a column index is within a table's column count.</summary>
    public static void ValidateColumnIndex(int columnIndex, int columnCount, string tableName)
    {
        if (columnIndex < 0 || columnIndex >= columnCount)
            throw new ArgumentOutOfRangeException(nameof(columnIndex),
                $"Column index {columnIndex} is out of range. Table '{tableName}' has {columnCount} column(s), valid range: 0-{columnCount - 1}.");
    }
}
