using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ImageMagick;
using PptxMcp.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Services;

public partial class PresentationService
{
    /// <summary>
    /// Optimize images in a PPTX file by downscaling, converting formats, and recompressing.
    /// </summary>
    /// <param name="filePath">Path to the PPTX file.</param>
    /// <param name="targetDpi">Target DPI for display (default 150 for screen).</param>
    /// <param name="jpegQuality">JPEG quality 1-100 (default 85).</param>
    /// <param name="convertFormats">Convert BMP/TIFF to PNG/JPEG (default true).</param>
    public ImageOptimizationResult OptimizeImages(
        string filePath,
        int targetDpi = 150,
        int jpegQuality = 85,
        bool convertFormats = true)
    {
        using var doc = PresentationDocument.Open(filePath, true);
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation part is missing.");

        var validator = new OpenXmlValidator();
        int errorsBefore = validator.Validate(doc).Count();

        var optimizedImages = new List<OptimizedImageInfo>();
        long totalBytesBefore = 0;
        long totalBytesAfter = 0;

        // Collect all owner parts (slides, layouts, masters) and their images.
        var allOwnerParts = CollectAllOwnerParts(presentationPart);

        // Track which ImageParts we've already processed (same part may be shared).
        var processedImageUris = new HashSet<string>();

        foreach (var ownerPart in allOwnerParts)
        {
            foreach (var idPartPair in ownerPart.Parts)
            {
                if (idPartPair.OpenXmlPart is not ImagePart imagePart)
                    continue;

                var uri = imagePart.Uri.ToString();
                if (!processedImageUris.Add(uri))
                    continue; // Already processed this image

                var imageInfo = OptimizeImagePart(
                    ownerPart,
                    imagePart,
                    targetDpi,
                    jpegQuality,
                    convertFormats);

                if (imageInfo is not null)
                {
                    optimizedImages.Add(imageInfo);
                    totalBytesBefore += imageInfo.OriginalSizeBytes;
                    totalBytesAfter += imageInfo.OptimizedSizeBytes;
                }
            }
        }

        // Save and validate after modification.
        presentationPart.Presentation.Save();
        int errorsAfter = validator.Validate(doc).Count();

        int imagesProcessed = optimizedImages.Count(i => i.BytesSaved > 0);
        int imagesSkipped = optimizedImages.Count(i => i.BytesSaved == 0);
        long totalBytesSaved = totalBytesBefore - totalBytesAfter;

        string message = imagesProcessed > 0
            ? $"Optimized {imagesProcessed} image(s), skipped {imagesSkipped}. Saved {totalBytesSaved:N0} bytes."
            : "No images optimized.";

        return new ImageOptimizationResult(
            Success: true,
            FilePath: filePath,
            ImagesProcessed: imagesProcessed,
            ImagesSkipped: imagesSkipped,
            TotalBytesBefore: totalBytesBefore,
            TotalBytesAfter: totalBytesAfter,
            TotalBytesSaved: totalBytesSaved,
            OptimizedImages: optimizedImages,
            Validation: new ValidationStatus(errorsBefore, errorsAfter, errorsAfter == 0),
            Message: message);
    }

    /// <summary>
    /// Optimize a single image part based on display dimensions and target DPI.
    /// Returns optimization details or null if the image should be skipped.
    /// </summary>
    private static OptimizedImageInfo? OptimizeImagePart(
        OpenXmlPart ownerPart,
        ImagePart imagePart,
        int targetDpi,
        int jpegQuality,
        bool convertFormats)
    {
        var originalCopy = new MemoryStream();
        using (var originalStream = imagePart.GetStream())
        {
            originalStream.CopyTo(originalCopy);
        }
        originalCopy.Position = 0;

        long originalSize = originalCopy.Length;

        // Read image metadata with lightweight MagickImageInfo.
        MagickImageInfo info;
        try
        {
            info = new MagickImageInfo(originalCopy);
            originalCopy.Position = 0;
        }
        catch
        {
            // Corrupted or unsupported image format.
            return null;
        }

        int originalWidth = (int)info.Width;
        int originalHeight = (int)info.Height;
        var originalFormat = info.Format.ToString();

        // Find the display dimensions of this image on the slide.
        var displaySize = GetImageDisplaySize(ownerPart, imagePart);

        int targetWidth = originalWidth;
        int targetHeight = originalHeight;
        bool needsDownscaling = false;

        if (displaySize is not null)
        {
            // Convert EMU to pixels at target DPI.
            // 1 inch = 914400 EMU; pixels = emu / 914400 * dpi
            double displayWidthPixels = displaySize.Value.Cx / 914400.0 * targetDpi;
            double displayHeightPixels = displaySize.Value.Cy / 914400.0 * targetDpi;

            // Downscale if image is significantly larger than display size.
            if (originalWidth > displayWidthPixels * 1.1 || originalHeight > displayHeightPixels * 1.1)
            {
                needsDownscaling = true;
                double aspectRatio = (double)originalWidth / originalHeight;
                targetWidth = (int)Math.Ceiling(displayWidthPixels);
                targetHeight = (int)Math.Ceiling(displayHeightPixels);

                // Preserve aspect ratio.
                if (targetWidth / (double)targetHeight > aspectRatio)
                    targetWidth = (int)Math.Ceiling(targetHeight * aspectRatio);
                else
                    targetHeight = (int)Math.Ceiling(targetWidth / aspectRatio);
            }
        }

        // Determine if format conversion is needed.
        bool needsConversion = convertFormats &&
            (info.Format == MagickFormat.Bmp ||
             info.Format == MagickFormat.Tiff ||
             info.Format == MagickFormat.Tiff64);

        // Determine target format.
        MagickFormat targetFormat = info.Format;
        if (needsConversion)
        {
            // Convert BMP/TIFF to PNG for lossless, or JPEG for photos.
            // Use PNG as default for safety.
            targetFormat = MagickFormat.Png;
        }

        // Perform optimization if needed.
        if (!needsDownscaling && !needsConversion && info.Format != MagickFormat.Jpeg)
        {
            // No optimization possible.
            return new OptimizedImageInfo(
                ImagePath: imagePart.Uri.ToString(),
                OriginalFormat: originalFormat,
                OptimizedFormat: originalFormat,
                OriginalWidth: originalWidth,
                OriginalHeight: originalHeight,
                OptimizedWidth: originalWidth,
                OptimizedHeight: originalHeight,
                OriginalSizeBytes: originalSize,
                OptimizedSizeBytes: originalSize,
                BytesSaved: 0,
                Action: "skipped");
        }

        using var image = new MagickImage(originalCopy);

        bool modified = false;
        var actions = new List<string>();

        // Downscale if needed.
        if (needsDownscaling)
        {
            image.Resize((uint)targetWidth, (uint)targetHeight);
            actions.Add("downscaled");
            modified = true;
        }

        // Convert format if needed.
        if (needsConversion)
        {
            image.Format = targetFormat;
            actions.Add($"converted to {targetFormat}");
            modified = true;
        }

        // Recompress JPEG.
        if (image.Format == MagickFormat.Jpeg)
        {
            image.Quality = (uint)jpegQuality;
            actions.Add("recompressed");
            modified = true;
        }

        if (!modified)
        {
            return new OptimizedImageInfo(
                ImagePath: imagePart.Uri.ToString(),
                OriginalFormat: originalFormat,
                OptimizedFormat: originalFormat,
                OriginalWidth: originalWidth,
                OriginalHeight: originalHeight,
                OptimizedWidth: originalWidth,
                OptimizedHeight: originalHeight,
                OriginalSizeBytes: originalSize,
                OptimizedSizeBytes: originalSize,
                BytesSaved: 0,
                Action: "skipped");
        }

        // Write optimized image to memory.
        var optimizedStream = new MemoryStream();
        image.Write(optimizedStream);
        optimizedStream.Position = 0;

        long optimizedSize = optimizedStream.Length;

        // Only replace if the new image is smaller.
        if (optimizedSize < originalSize)
        {
            imagePart.FeedData(optimizedStream);

            return new OptimizedImageInfo(
                ImagePath: imagePart.Uri.ToString(),
                OriginalFormat: originalFormat,
                OptimizedFormat: image.Format.ToString(),
                OriginalWidth: originalWidth,
                OriginalHeight: originalHeight,
                OptimizedWidth: (int)image.Width,
                OptimizedHeight: (int)image.Height,
                OriginalSizeBytes: originalSize,
                OptimizedSizeBytes: optimizedSize,
                BytesSaved: originalSize - optimizedSize,
                Action: string.Join(", ", actions));
        }
        else
        {
            // Optimized image is larger; skip replacement.
            return new OptimizedImageInfo(
                ImagePath: imagePart.Uri.ToString(),
                OriginalFormat: originalFormat,
                OptimizedFormat: originalFormat,
                OriginalWidth: originalWidth,
                OriginalHeight: originalHeight,
                OptimizedWidth: originalWidth,
                OptimizedHeight: originalHeight,
                OriginalSizeBytes: originalSize,
                OptimizedSizeBytes: originalSize,
                BytesSaved: 0,
                Action: "skipped (no savings)");
        }
    }

    /// <summary>
    /// Get the display dimensions (in EMU) of an image on a slide.
    /// Returns null if the image is not directly displayed (e.g., embedded in layout/master).
    /// </summary>
    private static (long Cx, long Cy)? GetImageDisplaySize(OpenXmlPart ownerPart, ImagePart imagePart)
    {
        if (ownerPart is not SlidePart slidePart)
            return null; // Only process images directly on slides.

        var relId = FindRelationshipId(ownerPart, imagePart);
        if (relId is null)
            return null;

        // Find the Picture shape that references this image via Blip.Embed.
        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree is null)
            return null;

        foreach (var picture in shapeTree.Elements<P.Picture>())
        {
            var blip = picture.GetFirstChild<P.BlipFill>()?.GetFirstChild<A.Blip>();
            if (blip?.Embed?.Value == relId)
            {
                // Found the picture shape; extract display dimensions.
                var extents = picture.ShapeProperties?.Transform2D?.Extents;
                if (extents?.Cx?.HasValue == true && extents.Cy?.HasValue == true)
                {
                    return (extents.Cx.Value, extents.Cy.Value);
                }
            }
        }

        return null;
    }
}
