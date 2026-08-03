using System.Security.Cryptography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PptxTools.Models;

namespace PptxTools.Services;

public partial class PresentationService
{
    /// <summary>
    /// Deduplicate identical media parts in a PPTX file.
    /// Finds media with the same SHA256 hash, redirects all references to a single canonical copy,
    /// and removes the orphaned duplicates. Validates the package before and after modification.
    /// </summary>
    public DeduplicateMediaResult DeduplicateMedia(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation part is missing.");

        var validator = new OpenXmlValidator();
        int errorsBefore = validator.Validate(doc).Count();

        // Phase 1: Build hash→ImageParts map across all owning parts.
        var hashToImageParts = new Dictionary<string, List<(ImagePart Part, string Uri)>>();
        var allOwnerParts = CollectAllOwnerParts(presentationPart);

        // Track which ImagePart URIs we've already hashed (same part referenced by multiple owners).
        var hashedUris = new Dictionary<string, string>(); // uri → hash

        foreach (var ownerPart in allOwnerParts)
        {
            foreach (var idPartPair in ownerPart.Parts)
            {
                if (idPartPair.OpenXmlPart is not ImagePart imagePart)
                    continue;

                var uri = imagePart.Uri.ToString();
                if (!hashedUris.TryGetValue(uri, out var hash))
                {
                    using var stream = imagePart.GetStream();
                    using var ms = new MemoryStream();
                    stream.CopyTo(ms);
                    ms.Position = 0;
                    hash = Convert.ToHexString(SHA256.HashData(ms));
                    hashedUris[uri] = hash;
                }

                if (!hashToImageParts.TryGetValue(hash, out var list))
                {
                    list = [];
                    hashToImageParts[hash] = list;
                }

                // Only add each URI once to the group.
                if (!list.Any(x => x.Uri == uri))
                    list.Add((imagePart, uri));
            }
        }

        // Phase 2: For each duplicate group, pick canonical and redirect references.
        var groups = new List<DeduplicatedGroupInfo>();
        int totalPartsRemoved = 0;
        long totalBytesSaved = 0;

        foreach (var (hash, parts) in hashToImageParts)
        {
            if (parts.Count < 2)
                continue;

            // Canonical = first alphabetically by URI.
            var sorted = parts.OrderBy(p => p.Uri, StringComparer.OrdinalIgnoreCase).ToList();
            var canonical = sorted[0];
            var duplicates = sorted.Skip(1).ToList();

            long sizePerCopy = 0;
            try
            {
                using var s = canonical.Part.GetStream();
                sizePerCopy = s.Length;
            }
            catch { }

            int referencesUpdated = 0;
            var removedUris = new List<string>();

            foreach (var dup in duplicates)
            {
                // Collect owners that reference this duplicate before modifying anything.
                var ownersWithDup = new List<(OpenXmlPart Owner, string OldRelId)>();
                foreach (var ownerPart in allOwnerParts)
                {
                    var oldRelId = FindRelationshipId(ownerPart, dup.Part);
                    if (oldRelId is not null)
                        ownersWithDup.Add((ownerPart, oldRelId));
                }

                // Redirect each owner's references from duplicate to canonical.
                foreach (var (ownerPart, oldRelId) in ownersWithDup)
                {
                    var newRelId = FindRelationshipId(ownerPart, canonical.Part)
                        ?? ownerPart.CreateRelationshipToPart(canonical.Part);

                    referencesUpdated += UpdateBlipReferences(ownerPart, oldRelId, newRelId);
                }

                // Now remove the duplicate part from each owner that had it.
                // DeletePart removes the relationship; the part is removed when the last ref goes.
                foreach (var (ownerPart, _) in ownersWithDup)
                {
                    if (ownerPart.Parts.Any(p => p.OpenXmlPart == dup.Part))
                        ownerPart.DeletePart(dup.Part);
                }

                removedUris.Add(dup.Uri);
                totalBytesSaved += sizePerCopy;
                totalPartsRemoved++;
            }

            groups.Add(new DeduplicatedGroupInfo(
                Hash: hash,
                ContentType: canonical.Part.ContentType,
                CanonicalPartUri: canonical.Uri,
                RemovedPartUris: removedUris.ToArray(),
                SizePerCopy: sizePerCopy,
                ReferencesUpdated: referencesUpdated));
        }

        // Phase 4: Save and validate after modification.
        presentationPart.Presentation.Save();
        int errorsAfter = validator.Validate(doc).Count();

        string message = groups.Count > 0
            ? $"Deduplicated {groups.Count} group(s), removed {totalPartsRemoved} part(s). Saved approximately {totalBytesSaved:N0} bytes."
            : "No duplicate media found.";

        return new DeduplicateMediaResult(
            Success: true,
            FilePath: filePath,
            DuplicateGroupsFound: groups.Count,
            PartsRemoved: totalPartsRemoved,
            BytesSaved: totalBytesSaved,
            Groups: groups,
            Validation: new ValidationStatus(errorsBefore, errorsAfter, errorsAfter == 0),
            Message: message);
    }

    /// <summary>Collect all parts that may own media relationships (slides, layouts, masters).</summary>
    private static List<OpenXmlPart> CollectAllOwnerParts(PresentationPart presentationPart)
    {
        var owners = new List<OpenXmlPart>();

        var slideIds = presentationPart.Presentation.SlideIdList
            ?.Elements<DocumentFormat.OpenXml.Presentation.SlideId>() ?? [];
        foreach (var slideId in slideIds)
        {
            if (slideId.RelationshipId?.Value is { } relId &&
                presentationPart.GetPartById(relId) is OpenXmlPart slidePart)
            {
                owners.Add(slidePart);
            }
        }

        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            owners.Add(masterPart);
            foreach (var layoutPart in masterPart.SlideLayoutParts)
                owners.Add(layoutPart);
        }

        return owners;
    }

    /// <summary>Find the relationship ID that an owner part uses to reference a target part, or null.</summary>
    private static string? FindRelationshipId(OpenXmlPart ownerPart, OpenXmlPart targetPart)
    {
        foreach (var idPart in ownerPart.Parts)
        {
            if (idPart.OpenXmlPart == targetPart)
                return idPart.RelationshipId;
        }
        return null;
    }

    /// <summary>
    /// Update all Blip.Embed attributes in an owner part's XML tree from oldRelId to newRelId.
    /// Returns the number of references updated.
    /// </summary>
    private static int UpdateBlipReferences(OpenXmlPart ownerPart, string oldRelId, string newRelId)
    {
        var rootElement = ownerPart.RootElement;
        if (rootElement is null)
            return 0;

        int count = 0;
        foreach (var blip in rootElement.Descendants<Blip>())
        {
            if (blip.Embed?.Value == oldRelId)
            {
                blip.Embed = newRelId;
                count++;
            }
        }

        if (count > 0)
            rootElement.Save();

        return count;
    }
}
