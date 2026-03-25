namespace PptxTools.Models;

/// <summary>Information about a slide master in the presentation.</summary>
/// <param name="Name">Display name of the slide master.</param>
/// <param name="Uri">Package URI of the slide master part.</param>
/// <param name="SizeBytes">Total size in bytes of the master and its exclusive relationship parts.</param>
/// <param name="IsUsed">True when at least one layout under this master is referenced by a slide.</param>
/// <param name="LayoutCount">Total number of layouts belonging to this master.</param>
/// <param name="UsedLayoutCount">Number of layouts under this master that are referenced by at least one slide.</param>
public record MasterInfo(
    string Name,
    string Uri,
    long SizeBytes,
    bool IsUsed,
    int LayoutCount,
    int UsedLayoutCount);

/// <summary>Information about a slide layout in the presentation.</summary>
/// <param name="Name">Display name of the slide layout.</param>
/// <param name="Uri">Package URI of the slide layout part.</param>
/// <param name="SizeBytes">Total size in bytes of the layout and its exclusive relationship parts.</param>
/// <param name="IsUsed">True when this layout is referenced by at least one slide.</param>
/// <param name="MasterName">Name of the parent slide master.</param>
/// <param name="ReferencedBySlides">1-based slide numbers that reference this layout.</param>
public record LayoutInfo(
    string Name,
    string Uri,
    long SizeBytes,
    bool IsUsed,
    string MasterName,
    IReadOnlyList<int> ReferencedBySlides);

/// <summary>Structured result for pptx_find_unused_layouts.</summary>
/// <param name="Success">True when the analysis completed without errors.</param>
/// <param name="FilePath">Path to the analyzed presentation file.</param>
/// <param name="TotalMasters">Total number of slide masters in the presentation.</param>
/// <param name="TotalLayouts">Total number of slide layouts across all masters.</param>
/// <param name="UnusedMasterCount">Number of masters with no layouts referenced by any slide.</param>
/// <param name="UnusedLayoutCount">Number of layouts not referenced by any slide.</param>
/// <param name="EstimatedSavingsBytes">Estimated bytes recoverable by removing all unused masters and layouts.</param>
/// <param name="Masters">Details for each slide master.</param>
/// <param name="Layouts">Details for each slide layout.</param>
/// <param name="Warnings">Advisory messages such as orphan-layout risks.</param>
/// <param name="Message">Human-readable status message describing the outcome.</param>
public record UnusedLayoutsResult(
    bool Success,
    string FilePath,
    int TotalMasters,
    int TotalLayouts,
    int UnusedMasterCount,
    int UnusedLayoutCount,
    long EstimatedSavingsBytes,
    IReadOnlyList<MasterInfo> Masters,
    IReadOnlyList<LayoutInfo> Layouts,
    IReadOnlyList<string> Warnings,
    string Message);
