using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxMcp.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Services;

public partial class PresentationService
{
    private static readonly MethodInfo AddNewPartWithContentTypeAndIdMethod = typeof(OpenXmlPartContainer)
        .GetMethods(BindingFlags.Public | BindingFlags.Instance)
        .Single(method =>
            method.Name == nameof(OpenXmlPartContainer.AddNewPart)
            && method.IsGenericMethodDefinition
            && method.GetParameters() is var parameters
            && parameters.Length == 2
            && parameters[0].ParameterType == typeof(string)
            && parameters[1].ParameterType == typeof(string));

    public AddSlideFromLayoutResult AddSlideFromLayout(string filePath, string layoutName, IReadOnlyDictionary<string, string>? placeholderValues = null, int? insertAt = null)
    {
        if (string.IsNullOrWhiteSpace(layoutName))
            throw new ArgumentException("layoutName is required.", nameof(layoutName));

        using var doc = PresentationDocument.Open(filePath, true);
        var presentationPart = doc.PresentationPart!;
        var resolvedLayout = ResolveLayoutPart(presentationPart, layoutName);
        var requests = ParsePlaceholderRequests(placeholderValues);

        ValidatePlaceholderRequests(
            GetPlaceholderTargets(resolvedLayout.SlideLayoutPart.SlideLayout.CommonSlideData?.ShapeTree),
            requests,
            $"layout '{resolvedLayout.LayoutName}'");

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.Slide = CreateSlideFromLayout(resolvedLayout.SlideLayoutPart);
        slidePart.AddPart(resolvedLayout.SlideLayoutPart);

        var placeholdersPopulated = ApplyPlaceholderRequests(
            GetPlaceholderTargets(slidePart.Slide.CommonSlideData?.ShapeTree),
            requests,
            $"new slide from layout '{resolvedLayout.LayoutName}'");

        slidePart.Slide.Save();
        var slideNumber = InsertSlidePart(presentationPart, slidePart, insertAt);
        presentationPart.Presentation.Save();

        return new AddSlideFromLayoutResult(
            Success: true,
            SlideNumber: slideNumber,
            LayoutName: resolvedLayout.LayoutName,
            PlaceholdersPopulated: placeholdersPopulated,
            Message: $"Added slide {slideNumber} from layout '{resolvedLayout.LayoutName}'.");
    }

    public DuplicateSlideResult DuplicateSlide(string filePath, int slideNumber, IReadOnlyDictionary<string, string>? placeholderOverrides = null, int? insertAt = null)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var presentationPart = doc.PresentationPart!;
        var slideIds = GetSlideIds(doc);
        if (slideNumber < 1 || slideNumber > slideIds.Count)
            throw new ArgumentOutOfRangeException(nameof(slideNumber), $"slideNumber {slideNumber} is out of range. Presentation has {slideIds.Count} slide(s).");

        var sourceSlidePart = GetSlidePart(doc, slideIds, slideNumber - 1);
        var requests = ParsePlaceholderRequests(placeholderOverrides);

        ValidatePlaceholderRequests(
            GetPlaceholderTargets(sourceSlidePart.Slide.CommonSlideData?.ShapeTree),
            requests,
            $"slide {slideNumber}");

        var duplicatedSlidePart = presentationPart.AddNewPart<SlidePart>();
        var clonedParts = new Dictionary<OpenXmlPart, OpenXmlPart>(ReferenceEqualityComparer.Instance)
        {
            [sourceSlidePart] = duplicatedSlidePart
        };

        CopyPartContent(sourceSlidePart, duplicatedSlidePart);
        ClonePartRelationships(sourceSlidePart, duplicatedSlidePart, clonedParts);

        var overridesApplied = ApplyPlaceholderRequests(
            GetPlaceholderTargets(duplicatedSlidePart.Slide.CommonSlideData?.ShapeTree),
            requests,
            $"duplicated slide {slideNumber}");

        duplicatedSlidePart.Slide.Save();

        var targetSlideNumber = insertAt ?? (slideNumber + 1);
        var newSlideNumber = InsertSlidePart(presentationPart, duplicatedSlidePart, targetSlideNumber);
        presentationPart.Presentation.Save();

        return new DuplicateSlideResult(
            Success: true,
            NewSlideNumber: newSlideNumber,
            ShapesCopied: CountRenderableShapes(duplicatedSlidePart.Slide),
            OverridesApplied: overridesApplied,
            Message: $"Duplicated slide {slideNumber} to slide {newSlideNumber}.");
    }

    private static ResolvedLayoutPart ResolveLayoutPart(PresentationPart presentationPart, string layoutName)
    {
        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            foreach (var layoutPart in masterPart.SlideLayoutParts)
            {
                var candidateName = GetLayoutName(layoutPart);
                if (string.Equals(candidateName, layoutName, StringComparison.OrdinalIgnoreCase))
                    return new ResolvedLayoutPart(layoutPart, candidateName);
            }
        }

        var availableLayouts = presentationPart.SlideMasterParts
            .SelectMany(masterPart => masterPart.SlideLayoutParts)
            .Select(GetLayoutName)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        throw new InvalidOperationException(
            availableLayouts.Count == 0
                ? "Presentation does not contain any slide layouts."
                : $"Layout '{layoutName}' was not found. Available layouts: {string.Join(", ", availableLayouts)}");
    }

    private static string GetLayoutName(SlideLayoutPart layoutPart) =>
        layoutPart.SlideLayout.CommonSlideData?.Name?.Value
        ?? layoutPart.SlideLayout.Type?.Value.ToString()
        ?? "Unnamed Layout";

    private static Slide CreateSlideFromLayout(SlideLayoutPart layoutPart)
    {
        var layoutShapeTree = layoutPart.SlideLayout.CommonSlideData?.ShapeTree;
        var shapeTree = layoutShapeTree is null
            ? CreateDefaultShapeTree()
            : CreateShapeTreeFromLayout(layoutShapeTree);

        return new Slide(
            new CommonSlideData(shapeTree),
            new ColorMapOverride(new A.MasterColorMapping()));
    }

    private static ShapeTree CreateShapeTreeFromLayout(ShapeTree layoutShapeTree)
    {
        var shapeTree = new ShapeTree();

        var groupProperties = layoutShapeTree.GetFirstChild<P.NonVisualGroupShapeProperties>();
        shapeTree.Append(groupProperties is null
            ? new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties())
            : (P.NonVisualGroupShapeProperties)groupProperties.CloneNode(true));

        var shapeProperties = layoutShapeTree.GetFirstChild<GroupShapeProperties>();
        shapeTree.Append(shapeProperties is null
            ? new GroupShapeProperties(new A.TransformGroup())
            : (GroupShapeProperties)shapeProperties.CloneNode(true));

        foreach (var child in layoutShapeTree.ChildElements)
        {
            if (GetPlaceholderShape(child) is null)
                continue;

            var clonedChild = (OpenXmlElement)child.CloneNode(true);
            ClearPlaceholderContent(clonedChild);
            shapeTree.Append(clonedChild);
        }

        return shapeTree;
    }

    private static ShapeTree CreateDefaultShapeTree() =>
        new(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

    private static void ClearPlaceholderContent(OpenXmlElement element)
    {
        if (element is Shape shape)
            ReplaceShapeTextPreservingFormatting(shape, string.Empty);
    }

    private static int InsertSlidePart(PresentationPart presentationPart, SlidePart slidePart, int? insertAt)
    {
        var slideIdList = presentationPart.Presentation.SlideIdList ??= new SlideIdList();
        var existingSlideIds = slideIdList.Elements<SlideId>().ToList();
        var slideNumber = insertAt ?? (existingSlideIds.Count + 1);

        if (slideNumber < 1 || slideNumber > existingSlideIds.Count + 1)
            throw new ArgumentOutOfRangeException(nameof(insertAt), $"insertAt {slideNumber} is out of range. Presentation has {existingSlideIds.Count} slide(s).");

        uint maxId = existingSlideIds.Count > 0
            ? existingSlideIds.Max(slideId => slideId.Id!.Value)
            : 255U;

        var newSlideId = new SlideId
        {
            Id = maxId + 1,
            RelationshipId = presentationPart.GetIdOfPart(slidePart)
        };

        if (slideNumber == existingSlideIds.Count + 1)
            slideIdList.Append(newSlideId);
        else
            slideIdList.InsertBefore(newSlideId, existingSlideIds[slideNumber - 1]);

        return slideNumber;
    }

    private static IReadOnlyList<PlaceholderRequest> ParsePlaceholderRequests(IReadOnlyDictionary<string, string>? placeholderValues)
    {
        if (placeholderValues is null || placeholderValues.Count == 0)
            return [];

        return placeholderValues
            .Select(entry => ParsePlaceholderRequest(entry.Key, entry.Value))
            .ToList();
    }

    private static PlaceholderRequest ParsePlaceholderRequest(string key, string value)
    {
        if (string.IsNullOrWhiteSpace(key))
            throw new ArgumentException("Placeholder keys cannot be empty.", nameof(key));

        var trimmedKey = key.Trim();
        var separatorIndex = trimmedKey.IndexOf(':');
        var typeToken = separatorIndex >= 0 ? trimmedKey[..separatorIndex] : trimmedKey;
        var indexToken = separatorIndex >= 0 ? trimmedKey[(separatorIndex + 1)..] : null;

        uint? placeholderIndex = null;
        if (!string.IsNullOrWhiteSpace(indexToken))
        {
            if (!uint.TryParse(indexToken, out var parsedIndex))
                throw new ArgumentException($"Placeholder key '{trimmedKey}' has an invalid index. Use keys like 'Title' or 'Body:1'.", nameof(key));

            placeholderIndex = parsedIndex;
        }

        return new PlaceholderRequest(trimmedKey, NormalizePlaceholderType(typeToken), placeholderIndex, value);
    }

    private static string NormalizePlaceholderType(string typeToken)
    {
        if (string.IsNullOrWhiteSpace(typeToken))
            throw new ArgumentException("Placeholder type is required.", nameof(typeToken));

        var normalized = new string(typeToken.Where(character => !char.IsWhiteSpace(character) && character != '_' && character != '-').ToArray()).ToLowerInvariant();
        return normalized switch
        {
            "title" or "centeredtitle" or "ctrtitle" => "Title",
            "subtitle" => "SubTitle",
            "body" => "Body",
            "picture" or "image" => "Picture",
            "object" or "content" => "Object",
            "chart" => "Chart",
            "table" => "Table",
            "media" => "Media",
            "clipart" => "ClipArt",
            _ => throw new ArgumentException($"Unsupported placeholder type '{typeToken}'.", nameof(typeToken))
        };
    }

    private static void ValidatePlaceholderRequests(IReadOnlyList<PlaceholderTarget> targets, IReadOnlyList<PlaceholderRequest> requests, string context)
    {
        if (requests.Count == 0)
            return;

        ResolvePlaceholderAssignments(targets, requests, context);
    }

    private static int ApplyPlaceholderRequests(IReadOnlyList<PlaceholderTarget> targets, IReadOnlyList<PlaceholderRequest> requests, string context)
    {
        if (requests.Count == 0)
            return 0;

        var assignments = ResolvePlaceholderAssignments(targets, requests, context);
        foreach (var assignment in assignments)
            ReplaceShapeTextPreservingFormatting(assignment.Target.Shape!, assignment.Request.Value);

        return assignments.Count;
    }

    private static IReadOnlyList<PlaceholderAssignment> ResolvePlaceholderAssignments(IReadOnlyList<PlaceholderTarget> targets, IReadOnlyList<PlaceholderRequest> requests, string context)
    {
        var assignments = new List<PlaceholderAssignment>(requests.Count);

        foreach (var request in requests)
        {
            var matches = targets
                .Where(target => string.Equals(target.SemanticType, request.SemanticType, StringComparison.OrdinalIgnoreCase))
                .Where(target => request.PlaceholderIndex is null || target.PlaceholderIndex == request.PlaceholderIndex)
                .OrderBy(target => target.Order)
                .ToList();

            if (matches.Count == 0)
                throw new InvalidOperationException($"Placeholder '{request.Key}' was not found on {context}. Available placeholders: {DescribePlaceholders(targets)}");

            if (request.PlaceholderIndex is null && matches.Count > 1)
                throw new InvalidOperationException($"Placeholder '{request.Key}' is ambiguous on {context}. Use 'Type:Index'. Available placeholders: {DescribePlaceholders(targets)}");

            var match = matches[0];
            if (match.Shape is null)
                throw new InvalidOperationException($"Placeholder '{request.Key}' on {context} is not text-capable.");

            assignments.Add(new PlaceholderAssignment(request, match));
        }

        var duplicateTarget = assignments
            .GroupBy(assignment => assignment.Target.Order)
            .FirstOrDefault(group => group.Count() > 1);

        if (duplicateTarget is not null)
        {
            var conflictingKeys = string.Join(", ", duplicateTarget.Select(assignment => assignment.Request.Key));
            throw new InvalidOperationException($"Multiple placeholder values target the same placeholder on {context}: {conflictingKeys}");
        }

        return assignments;
    }

    private static string DescribePlaceholders(IEnumerable<PlaceholderTarget> targets) =>
        string.Join(", ", targets.Select(target => target.PlaceholderIndex is null
            ? target.SemanticType
            : $"{target.SemanticType}:{target.PlaceholderIndex}"));

    private static IReadOnlyList<PlaceholderTarget> GetPlaceholderTargets(ShapeTree? shapeTree)
    {
        if (shapeTree is null)
            return [];

        return shapeTree.ChildElements
            .Select((child, index) => CreatePlaceholderTarget(child, index))
            .Where(target => target is not null)
            .Cast<PlaceholderTarget>()
            .ToList();
    }

    private static PlaceholderTarget? CreatePlaceholderTarget(OpenXmlElement element, int order)
    {
        var placeholderShape = GetPlaceholderShape(element);
        if (placeholderShape is null)
            return null;

        return new PlaceholderTarget(
            Order: order,
            SemanticType: GetSemanticPlaceholderType(placeholderShape),
            PlaceholderIndex: placeholderShape.Index?.Value,
            Shape: element as Shape);
    }

    private static string GetSemanticPlaceholderType(PlaceholderShape placeholderShape)
    {
        var placeholderType = placeholderShape.Type?.Value;
        if (placeholderType is null)
            return placeholderShape.Index?.Value == 0 ? "Title" : "Object";

        if (placeholderType == PlaceholderValues.Title || placeholderType == PlaceholderValues.CenteredTitle)
            return "Title";
        if (placeholderType == PlaceholderValues.SubTitle)
            return "SubTitle";
        if (placeholderType == PlaceholderValues.Body)
            return "Body";
        if (placeholderType == PlaceholderValues.Picture)
            return "Picture";
        if (placeholderType == PlaceholderValues.Object)
            return "Object";
        if (placeholderType == PlaceholderValues.Chart)
            return "Chart";
        if (placeholderType == PlaceholderValues.Table)
            return "Table";
        if (placeholderType == PlaceholderValues.Media)
            return "Media";
        if (placeholderType == PlaceholderValues.ClipArt)
            return "ClipArt";

        return placeholderType.ToString();
    }

    private static PlaceholderShape? GetPlaceholderShape(OpenXmlElement element) => element switch
    {
        Shape shape => shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
        Picture picture => picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
        P.GraphicFrame frame => frame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
        _ => null
    };

    private static int CountRenderableShapes(Slide slide)
    {
        var shapeTree = slide.CommonSlideData?.ShapeTree;
        if (shapeTree is null)
            return 0;

        var count = 0;
        foreach (var child in shapeTree.ChildElements)
        {
            if (ExtractShape(child) is not null)
                count++;
        }

        return count;
    }

    private static void CopyPartContent(OpenXmlPart sourcePart, OpenXmlPart destinationPart)
    {
        using var sourceStream = sourcePart.GetStream(FileMode.Open, FileAccess.Read);
        destinationPart.FeedData(sourceStream);
    }

    private static void ClonePartRelationships(OpenXmlPart sourcePart, OpenXmlPart destinationPart, IDictionary<OpenXmlPart, OpenXmlPart> clonedParts)
    {
        foreach (var childPart in sourcePart.Parts)
        {
            if (ShouldSharePartAcrossSlides(childPart.OpenXmlPart))
            {
                destinationPart.AddPart(childPart.OpenXmlPart, childPart.RelationshipId);
                continue;
            }

            if (!clonedParts.TryGetValue(childPart.OpenXmlPart, out var clonedChildPart))
            {
                clonedChildPart = AddNewPartLike(destinationPart, childPart.OpenXmlPart, childPart.RelationshipId);
                clonedParts[childPart.OpenXmlPart] = clonedChildPart;
                CopyPartContent(childPart.OpenXmlPart, clonedChildPart);
                ClonePartRelationships(childPart.OpenXmlPart, clonedChildPart, clonedParts);
            }
            else
            {
                destinationPart.CreateRelationshipToPart(clonedChildPart, childPart.RelationshipId);
            }
        }

        foreach (var externalRelationship in sourcePart.ExternalRelationships)
            destinationPart.AddExternalRelationship(externalRelationship.RelationshipType, externalRelationship.Uri, externalRelationship.Id);

        foreach (var hyperlinkRelationship in sourcePart.HyperlinkRelationships)
            destinationPart.AddHyperlinkRelationship(hyperlinkRelationship.Uri, hyperlinkRelationship.IsExternal, hyperlinkRelationship.Id);
    }

    private static bool ShouldSharePartAcrossSlides(OpenXmlPart part) =>
        part is SlideLayoutPart
        or SlideMasterPart
        or NotesMasterPart
        or ThemePart;

    private static OpenXmlPart AddNewPartLike(OpenXmlPartContainer container, OpenXmlPart sourcePart, string relationshipId) =>
        (OpenXmlPart)AddNewPartWithContentTypeAndIdMethod
            .MakeGenericMethod(sourcePart.GetType())
            .Invoke(container, [sourcePart.ContentType, relationshipId])!;

    private sealed record ResolvedLayoutPart(SlideLayoutPart SlideLayoutPart, string LayoutName);

    private sealed record PlaceholderRequest(string Key, string SemanticType, uint? PlaceholderIndex, string Value);

    private sealed record PlaceholderTarget(int Order, string SemanticType, uint? PlaceholderIndex, Shape? Shape);

    private sealed record PlaceholderAssignment(PlaceholderRequest Request, PlaceholderTarget Target);
}
