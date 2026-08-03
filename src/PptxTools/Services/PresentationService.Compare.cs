using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxTools.Models;

namespace PptxTools.Services;

public partial class PresentationService
{
    /// <summary>Compare two presentations and return a structured diff report.</summary>
    public ComparisonResult ComparePresentations(string sourcePath, string targetPath, CompareAction action)
    {
        var sourceSlides = GetAllSlideContents(sourcePath);
        var targetSlides = GetAllSlideContents(targetPath);

        var slideDiffs = action is CompareAction.MetadataOnly or CompareAction.TextOnly
            ? []
            : CompareSlideStructure(sourceSlides, targetSlides);

        var textDiffs = action is CompareAction.SlidesOnly or CompareAction.MetadataOnly
            ? []
            : CompareTextContent(sourceSlides, targetSlides);

        var metaDiffs = action is CompareAction.SlidesOnly or CompareAction.TextOnly
            ? []
            : CompareMetadataFields(sourcePath, targetPath);

        int totalDiffs = slideDiffs.Count + textDiffs.Count + metaDiffs.Count;
        bool identical = totalDiffs == 0;

        var message = identical
            ? "No differences found between the two presentations."
            : $"Found {totalDiffs} difference(s): {slideDiffs.Count} slide, {textDiffs.Count} text, {metaDiffs.Count} metadata.";

        return new ComparisonResult(
            Success: true,
            Action: action.ToString(),
            SourceFile: sourcePath,
            TargetFile: targetPath,
            AreIdentical: identical,
            DifferenceCount: totalDiffs,
            SlideDifferences: slideDiffs,
            TextDifferences: textDiffs,
            MetadataDifferences: metaDiffs,
            Message: message);
    }

    // --- Slide structure comparison ---

    private static List<SlideDifference> CompareSlideStructure(
        IReadOnlyList<SlideContent> sourceSlides,
        IReadOnlyList<SlideContent> targetSlides)
    {
        var diffs = new List<SlideDifference>();

        if (targetSlides.Count > sourceSlides.Count)
        {
            for (int i = sourceSlides.Count; i < targetSlides.Count; i++)
            {
                diffs.Add(new SlideDifference(
                    SlideNumber: i + 1,
                    DifferenceType: "Added",
                    Description: $"Slide {i + 1} exists in target but not in source."));
            }
        }
        else if (sourceSlides.Count > targetSlides.Count)
        {
            for (int i = targetSlides.Count; i < sourceSlides.Count; i++)
            {
                diffs.Add(new SlideDifference(
                    SlideNumber: i + 1,
                    DifferenceType: "Removed",
                    Description: $"Slide {i + 1} exists in source but not in target."));
            }
        }

        return diffs;
    }

    // --- Text content comparison ---

    private static List<TextDifference> CompareTextContent(
        IReadOnlyList<SlideContent> sourceSlides,
        IReadOnlyList<SlideContent> targetSlides)
    {
        var diffs = new List<TextDifference>();
        int overlapping = Math.Min(sourceSlides.Count, targetSlides.Count);

        for (int i = 0; i < overlapping; i++)
        {
            int slideNum = i + 1;
            CompareSlideShapeText(slideNum, sourceSlides[i], targetSlides[i], diffs);
        }

        return diffs;
    }

    private static void CompareSlideShapeText(
        int slideNumber, SlideContent source, SlideContent target,
        List<TextDifference> diffs)
    {
        var sourceShapes = source.Shapes
            .Where(s => s.Text is not null)
            .ToDictionary(s => s.Name, s => s);
        var targetShapes = target.Shapes
            .Where(s => s.Text is not null)
            .ToDictionary(s => s.Name, s => s);

        foreach (var (name, srcShape) in sourceShapes)
        {
            if (targetShapes.TryGetValue(name, out var tgtShape))
            {
                if (!string.Equals(srcShape.Text, tgtShape.Text, StringComparison.Ordinal))
                {
                    diffs.Add(new TextDifference(slideNumber, name, "Modified", srcShape.Text, tgtShape.Text));
                }
            }
            else
            {
                diffs.Add(new TextDifference(slideNumber, name, "Removed", srcShape.Text, null));
            }
        }

        foreach (var (name, tgtShape) in targetShapes)
        {
            if (!sourceShapes.ContainsKey(name))
            {
                diffs.Add(new TextDifference(slideNumber, name, "Added", null, tgtShape.Text));
            }
        }
    }

    // --- Metadata comparison ---

    private List<MetadataDifference> CompareMetadataFields(string sourcePath, string targetPath)
    {
        var src = GetPresentationMetadata(sourcePath);
        var tgt = GetPresentationMetadata(targetPath);
        var diffs = new List<MetadataDifference>();

        AddIfDifferent(diffs, "Title", src.Title, tgt.Title);
        AddIfDifferent(diffs, "Creator", src.Creator, tgt.Creator);
        AddIfDifferent(diffs, "Subject", src.Subject, tgt.Subject);
        AddIfDifferent(diffs, "Keywords", src.Keywords, tgt.Keywords);
        AddIfDifferent(diffs, "Description", src.Description, tgt.Description);
        AddIfDifferent(diffs, "LastModifiedBy", src.LastModifiedBy, tgt.LastModifiedBy);
        AddIfDifferent(diffs, "Category", src.Category, tgt.Category);
        AddIfDifferent(diffs, "Created", src.Created?.ToString("o"), tgt.Created?.ToString("o"));
        AddIfDifferent(diffs, "Modified", src.Modified?.ToString("o"), tgt.Modified?.ToString("o"));

        return diffs;
    }

    private static void AddIfDifferent(List<MetadataDifference> diffs, string property, string? source, string? target)
    {
        if (!string.Equals(source, target, StringComparison.Ordinal))
        {
            diffs.Add(new MetadataDifference(property, source, target));
        }
    }
}
