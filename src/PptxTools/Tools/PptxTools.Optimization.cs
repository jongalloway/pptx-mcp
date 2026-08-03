using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using PptxTools.Models;

namespace PptxTools.Tools;

public partial class PptxTools
{
    /// <summary>
    /// Analyze the file size breakdown of a PowerPoint presentation by category.
    /// Scans all parts in the PPTX package (ZIP structure) and reports sizes broken down into:
    /// slides, images, video/audio, slide masters, slide layouts, and other parts.
    /// Each category includes a subtotal and per-part detail (relative path, content type, size in bytes).
    /// The root level includes actual file size on disk and total uncompressed part size.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    [McpServerTool(Title = "Analyze File Size", ReadOnly = true, Idempotent = true)]
    public partial Task<string> pptx_analyze_file_size(string filePath) =>
        ExecuteToolStructured(filePath,
            () => _service.AnalyzeFileSize(filePath),
            error => new FileSizeAnalysisResult(
                Success: false,
                FilePath: filePath,
                TotalFileSize: 0,
                TotalPartSize: 0,
                Categories: EmptyFileSizeCategories,
                Message: error));

    private static readonly IReadOnlyList<FileSizeCategory> EmptyFileSizeCategories =
    [
        new("slides", 0, 0, []),
        new("images", 0, 0, []),
        new("video_audio", 0, 0, []),
        new("masters", 0, 0, []),
        new("layouts", 0, 0, []),
        new("other", 0, 0, []),
    ];

    /// <summary>
    /// Manage slide layouts in a PowerPoint presentation.
    /// Available actions:
    /// - Find: Identify unused layouts and masters with estimated space savings (read-only).
    /// - Remove: Remove unused layouts and orphaned masters, with OpenXML validation before and after.
    /// - SetType: Set the semantic type attribute (e.g. "title", "obj", "secHead") on a named layout.
    /// - ModifyPlaceholder: Update placeholder type/index on a layout placeholder by ordinal.
    /// - AddPlaceholder: Add a placeholder shape to a layout with explicit bounds and placeholder metadata.
    /// Natural workflow: Find (diagnostic) → Remove (action) or SetType/ModifyPlaceholder/AddPlaceholder (correction).
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file.</param>
    /// <param name="action">The layout management operation to perform: Find, Remove, SetType, ModifyPlaceholder, or AddPlaceholder.</param>
    /// <param name="layoutUris">Optional array of layout URIs to remove. Only used with Remove action. Omit to auto-detect all unused layouts.</param>
    /// <param name="layoutName">Display name of the layout to update. Required for SetType action.</param>
    /// <param name="layoutType">
    /// Semantic type to assign to the layout. Required for SetType action.
    /// Common values: title, tx, obj, twoObj, secHead, blank, picTx, tbl, chart, titleOnly, cust.
    /// </param>
    /// <param name="placeholderIndex">0-based placeholder ordinal to modify. Required for ModifyPlaceholder action.</param>
    /// <param name="newType">New placeholder type (e.g. title, ctrTitle, subTitle, body, pic, obj, chart, ftr, dt, sldNum). Required for ModifyPlaceholder action.</param>
    /// <param name="newIdx">Optional replacement placeholder idx value. Used by ModifyPlaceholder action.</param>
    /// <param name="type">Placeholder type to add. Required for AddPlaceholder action.</param>
    /// <param name="idx">Placeholder idx value to add. Required for AddPlaceholder action.</param>
    /// <param name="x">Placeholder X position in EMU. Required for AddPlaceholder action.</param>
    /// <param name="y">Placeholder Y position in EMU. Required for AddPlaceholder action.</param>
    /// <param name="cx">Placeholder width in EMU. Required for AddPlaceholder action.</param>
    /// <param name="cy">Placeholder height in EMU. Required for AddPlaceholder action.</param>
    [McpServerTool(Title = "Manage Layouts")]
    [McpMeta("consolidatedTool", true)]
    [McpMeta("actions", JsonValue = """["Find","Remove","SetType","ModifyPlaceholder","AddPlaceholder"]""")]
    public partial Task<string> pptx_manage_layouts(
        string filePath,
        ManageLayoutsAction action,
        string[]? layoutUris = null,
        string? layoutName = null,
        string? layoutType = null,
        int? placeholderIndex = null,
        string? newType = null,
        int? newIdx = null,
        string? type = null,
        int? idx = null,
        long? x = null,
        long? y = null,
        long? cx = null,
        long? cy = null)
    {
        return action switch
        {
            ManageLayoutsAction.Find => ExecuteToolStructured(filePath,
                () => _service.FindUnusedLayouts(filePath),
                error => new UnusedLayoutsResult(
                    Success: false,
                    FilePath: filePath,
                    TotalMasters: 0,
                    TotalLayouts: 0,
                    UnusedMasterCount: 0,
                    UnusedLayoutCount: 0,
                    EstimatedSavingsBytes: 0,
                    Masters: [],
                    Layouts: [],
                    Warnings: [],
                    Message: error)),

            ManageLayoutsAction.Remove => ExecuteToolStructured(filePath,
                () => _service.RemoveUnusedLayouts(filePath, layoutUris),
                error => new RemoveLayoutsResult(
                    Success: false,
                    FilePath: filePath,
                    RemovedItems: [],
                    LayoutsRemoved: 0,
                    MastersRemoved: 0,
                    BytesSaved: 0,
                    Validation: new ValidationStatus(0, 0, false),
                    Message: error)),

            ManageLayoutsAction.SetType => ExecuteToolStructured(filePath,
                () =>
                {
                    if (string.IsNullOrWhiteSpace(layoutName))
                        throw new ArgumentException("layoutName is required for the SetType action.");
                    if (string.IsNullOrWhiteSpace(layoutType))
                        throw new ArgumentException("layoutType is required for the SetType action.");
                    return _service.SetLayoutType(filePath, layoutName, layoutType);
                },
                error => new SetLayoutTypeResult(
                    Success: false,
                    FilePath: filePath,
                    LayoutName: layoutName,
                    LayoutType: layoutType,
                    Message: error)),

            ManageLayoutsAction.ModifyPlaceholder => ExecuteToolStructured(filePath,
                () =>
                {
                    if (string.IsNullOrWhiteSpace(layoutName))
                        throw new ArgumentException("layoutName is required for the ModifyPlaceholder action.");
                    if (placeholderIndex is null)
                        throw new ArgumentException("placeholderIndex is required for the ModifyPlaceholder action.");
                    if (string.IsNullOrWhiteSpace(newType))
                        throw new ArgumentException("newType is required for the ModifyPlaceholder action.");
                    return _service.ModifyLayoutPlaceholder(filePath, layoutName, placeholderIndex.Value, newType, newIdx);
                },
                error => new ModifyLayoutPlaceholderResult(
                    Success: false,
                    FilePath: filePath,
                    LayoutName: layoutName,
                    PlaceholderIndex: placeholderIndex,
                    NewType: newType,
                    NewIdx: newIdx,
                    Message: error)),

            ManageLayoutsAction.AddPlaceholder => ExecuteToolStructured(filePath,
                () =>
                {
                    if (string.IsNullOrWhiteSpace(layoutName))
                        throw new ArgumentException("layoutName is required for the AddPlaceholder action.");
                    if (string.IsNullOrWhiteSpace(type))
                        throw new ArgumentException("type is required for the AddPlaceholder action.");
                    if (idx is null)
                        throw new ArgumentException("idx is required for the AddPlaceholder action.");
                    if (x is null || y is null || cx is null || cy is null)
                        throw new ArgumentException("x, y, cx, and cy are required for the AddPlaceholder action.");
                    return _service.AddLayoutPlaceholder(filePath, layoutName, type, idx.Value, x.Value, y.Value, cx.Value, cy.Value);
                },
                error => new AddLayoutPlaceholderResult(
                    Success: false,
                    FilePath: filePath,
                    LayoutName: layoutName,
                    Type: type,
                    Idx: idx,
                    X: x,
                    Y: y,
                    Cx: cx,
                    Cy: cy,
                    Message: error)),

            _ => Task.FromResult(JsonSerializer.Serialize(
                new { Success = false, Message = $"Unknown action: {action}. Valid actions: Find, Remove, SetType, ModifyPlaceholder, AddPlaceholder." },
                IndentedJson))
        };
    }

    /// <summary>
    /// Optimize images in a PowerPoint presentation by downscaling, converting formats, and recompressing.
    /// Scans all images across slides, layouts, and masters. Downscales images that are larger than their
    /// display dimensions warrant based on target DPI. Converts BMP/TIFF to PNG/JPEG. Recompresses JPEG images
    /// at the specified quality level. Only replaces images when optimization results in smaller file size.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the .pptx file to modify.</param>
    /// <param name="targetDpi">Target DPI for screen display (default 150; use 300 for print).</param>
    /// <param name="jpegQuality">JPEG compression quality 1-100 (default 85; higher = larger file).</param>
    /// <param name="convertFormats">Convert BMP/TIFF to PNG/JPEG (default true).</param>
    [McpServerTool(Title = "Optimize Images")]
    public partial Task<string> pptx_optimize_images(
        string filePath,
        int targetDpi = 150,
        int jpegQuality = 85,
        bool convertFormats = true) =>
        ExecuteToolStructured(filePath,
            () => _service.OptimizeImages(filePath, targetDpi, jpegQuality, convertFormats),
            error => new ImageOptimizationResult(
                Success: false,
                FilePath: filePath,
                ImagesProcessed: 0,
                ImagesSkipped: 0,
                TotalBytesBefore: 0,
                TotalBytesAfter: 0,
                TotalBytesSaved: 0,
                OptimizedImages: [],
                Validation: new ValidationStatus(0, 0, false),
                Message: error));
}
