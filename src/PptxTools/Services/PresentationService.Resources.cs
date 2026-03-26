using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxTools.Models;

namespace PptxTools.Services;

public partial class PresentationService
{
    /// <summary>
    /// Extract metadata for every embedded image across all slides.
    /// Returns shape name, content type, relationship ID, and dimensions.
    /// </summary>
    public IReadOnlyList<ImageInfo> GetImageInfos(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var presentationPart = doc.PresentationPart;
        if (presentationPart is null) return [];

        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<SlideId>() ?? [];
        var result = new List<ImageInfo>();
        int slideNumber = 1;

        foreach (var slideId in slideIds)
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!);
            var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
            if (shapeTree is null) { slideNumber++; continue; }

            foreach (var picture in shapeTree.Descendants<Picture>())
            {
                var drawingProps = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
                var shapeName = drawingProps?.Name?.Value ?? "";

                var embed = picture.BlipFill?.Blip?.Embed?.Value;
                if (embed is null) continue;

                string contentType = "";
                string imageFormat = "";
                try
                {
                    var imagePart = (ImagePart)slidePart.GetPartById(embed);
                    contentType = imagePart.ContentType;
                    imageFormat = contentType.Split('/').LastOrDefault()?.ToUpperInvariant() ?? "";
                }
                catch
                {
                    // Relationship may not resolve to an ImagePart (e.g. linked images)
                }

                var xfrm = picture.ShapeProperties?.Transform2D;

                result.Add(new ImageInfo(
                    SlideNumber: slideNumber,
                    ShapeName: shapeName,
                    ContentType: contentType,
                    ImageFormat: imageFormat,
                    RelationshipId: embed,
                    WidthEmu: xfrm?.Extents?.Cx?.Value,
                    HeightEmu: xfrm?.Extents?.Cy?.Value));
            }

            slideNumber++;
        }

        return result;
    }

    /// <summary>
    /// Read presentation-level metadata from package properties.
    /// </summary>
    public PresentationMetadata GetPresentationMetadata(string filePath)
    {
        using var doc = PresentationDocument.Open(filePath, false);
        var props = doc.PackageProperties;
        var slideCount = doc.PresentationPart?.Presentation.SlideIdList?.Elements<SlideId>().Count() ?? 0;

        return new PresentationMetadata(
            Title: props.Title,
            Creator: props.Creator,
            Created: props.Created,
            Modified: props.Modified,
            Subject: props.Subject,
            Keywords: props.Keywords,
            Description: props.Description,
            LastModifiedBy: props.LastModifiedBy,
            Category: props.Category,
            SlideCount: slideCount);
    }
}
