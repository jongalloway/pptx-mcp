namespace PptxMcp.Models;

/// <summary>Information about a removed layout or master.</summary>
/// <param name="Name">Display name of the removed item.</param>
/// <param name="Uri">Package URI of the removed part.</param>
/// <param name="Type">Either "layout" or "master".</param>
/// <param name="SizeBytes">Size in bytes of the removed part and its exclusive relationships.</param>
public record RemovedItemInfo(
    string Name,
    string Uri,
    string Type,
    long SizeBytes);

/// <summary>OpenXML validation status captured before and after removal.</summary>
/// <param name="ErrorsBefore">Count of validation errors before removal.</param>
/// <param name="ErrorsAfter">Count of validation errors after removal.</param>
/// <param name="IsValid">True when ErrorsAfter is zero.</param>
public record ValidationStatus(
    int ErrorsBefore,
    int ErrorsAfter,
    bool IsValid);

/// <summary>Structured result for pptx_remove_unused_layouts.</summary>
/// <param name="Success">True when the operation completed without errors.</param>
/// <param name="FilePath">Path to the modified presentation file.</param>
/// <param name="RemovedItems">Details for each removed layout and master.</param>
/// <param name="LayoutsRemoved">Number of layouts removed.</param>
/// <param name="MastersRemoved">Number of masters removed.</param>
/// <param name="BytesSaved">Total bytes saved by removing unused parts.</param>
/// <param name="Validation">OpenXML validation status before and after.</param>
/// <param name="Message">Human-readable status or error message.</param>
public record RemoveLayoutsResult(
    bool Success,
    string FilePath,
    IReadOnlyList<RemovedItemInfo> RemovedItems,
    int LayoutsRemoved,
    int MastersRemoved,
    long BytesSaved,
    ValidationStatus Validation,
    string Message);
