namespace PptxTools.Models;

/// <summary>Actions for the consolidated pptx_validate_presentation tool.</summary>
public enum ValidationAction
{
    /// <summary>Run all validation checks on the presentation.</summary>
    Validate
}

/// <summary>Severity level for a validation issue.</summary>
public enum ValidationSeverity
{
    /// <summary>Structural problem that may cause data loss or corruption.</summary>
    Error,

    /// <summary>Potential problem that should be reviewed.</summary>
    Warning,

    /// <summary>Informational finding that does not indicate a problem.</summary>
    Info
}

/// <summary>A single validation issue detected in the presentation.</summary>
/// <param name="SlideNumber">1-based slide number where the issue was found, or null for presentation-level issues.</param>
/// <param name="Severity">Severity of the issue.</param>
/// <param name="Category">Classification of the issue (e.g. "DuplicateShapeId", "MissingImageReference").</param>
/// <param name="Description">Human-readable description of what was found.</param>
/// <param name="Recommendation">Suggested action to resolve the issue.</param>
/// <param name="XmlContext">XML element path or line/position reference for debugging (e.g. "p:spTree/p:sp[@id=2]" or "Line 4, Position 12").</param>
public record ValidationIssue(
    int? SlideNumber,
    ValidationSeverity Severity,
    string Category,
    string Description,
    string Recommendation,
    string? XmlContext = null);

/// <summary>Result of a presentation validation operation.</summary>
/// <param name="Success">True when validation completed without internal errors.</param>
/// <param name="Action">The action that was performed.</param>
/// <param name="IssueCount">Total number of issues detected.</param>
/// <param name="ErrorCount">Number of Error-severity issues.</param>
/// <param name="WarningCount">Number of Warning-severity issues.</param>
/// <param name="InfoCount">Number of Info-severity issues.</param>
/// <param name="Issues">All detected issues, grouped by severity.</param>
/// <param name="Message">Human-readable summary of the validation results.</param>
public record ValidationResult(
    bool Success,
    string Action,
    int IssueCount,
    int ErrorCount,
    int WarningCount,
    int InfoCount,
    IReadOnlyList<ValidationIssue> Issues,
    string Message);
