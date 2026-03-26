namespace PptxTools.Models;

/// <summary>Metadata for a single image embedded in a presentation slide.</summary>
/// <param name="SlideNumber">1-based slide number containing this image.</param>
/// <param name="ShapeName">Name of the picture shape from drawing properties.</param>
/// <param name="ContentType">MIME content type of the image (e.g. image/png).</param>
/// <param name="ImageFormat">Short format label derived from content type (e.g. PNG, JPEG).</param>
/// <param name="RelationshipId">OpenXML relationship ID linking the shape to the image part.</param>
/// <param name="WidthEmu">Image width in EMUs. Null when transform data is unavailable.</param>
/// <param name="HeightEmu">Image height in EMUs. Null when transform data is unavailable.</param>
public record ImageInfo(
    int SlideNumber,
    string ShapeName,
    string ContentType,
    string ImageFormat,
    string RelationshipId,
    long? WidthEmu,
    long? HeightEmu);
