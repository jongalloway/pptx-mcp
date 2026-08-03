namespace PptxTools.Models;

/// <summary>Structured result for pptx_deduplicate_media.</summary>
/// <param name="Success">True when deduplication completed successfully.</param>
/// <param name="FilePath">Path to the modified presentation file.</param>
/// <param name="DuplicateGroupsFound">Number of groups containing identical media (same SHA256 hash).</param>
/// <param name="PartsRemoved">Total number of duplicate media parts removed.</param>
/// <param name="BytesSaved">Total bytes saved by removing duplicate parts.</param>
/// <param name="Groups">Details for each deduplicated group.</param>
/// <param name="Validation">OpenXML validation status before and after.</param>
/// <param name="Message">Human-readable status or error message.</param>
public record DeduplicateMediaResult(
    bool Success,
    string FilePath,
    int DuplicateGroupsFound,
    int PartsRemoved,
    long BytesSaved,
    IReadOnlyList<DeduplicatedGroupInfo> Groups,
    ValidationStatus Validation,
    string Message);

/// <summary>Details about a single group of deduplicated media.</summary>
/// <param name="Hash">SHA256 hex hash shared by all parts in this group.</param>
/// <param name="ContentType">MIME content type of the media.</param>
/// <param name="CanonicalPartUri">URI of the canonical (kept) part.</param>
/// <param name="RemovedPartUris">URIs of the removed duplicate parts.</param>
/// <param name="SizePerCopy">Size of each individual copy in bytes.</param>
/// <param name="ReferencesUpdated">Number of relationship references redirected to the canonical part.</param>
public record DeduplicatedGroupInfo(
    string Hash,
    string ContentType,
    string CanonicalPartUri,
    string[] RemovedPartUris,
    long SizePerCopy,
    int ReferencesUpdated);
