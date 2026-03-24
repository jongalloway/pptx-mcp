using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ImageMagick;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Tests.Services;

/// <summary>
/// Service-level tests for OptimizeImages (Issue #85 — Image optimization).
/// Validates image downscaling, format conversion, JPEG recompression, and result structure.
/// </summary>
[Trait("Category", "Unit")]
public class ImageOptimizationTests : PptxTestBase
{
    // ──────────────────────────────────────────────────────────
    //  1. No images — returns ImagesProcessed=0, Success=true
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void OptimizeImages_NoImages_ReturnsEmpty()
    {
        var path = CreateMinimalPptx("No Images");

        var result = Service.OptimizeImages(path);

        Assert.True(result.Success);
        Assert.Equal(0, result.ImagesProcessed);
        Assert.Equal(0, result.ImagesSkipped);
        Assert.Equal(0, result.TotalBytesBefore);
        Assert.Equal(0, result.TotalBytesAfter);
        Assert.Equal(0, result.TotalBytesSaved);
        Assert.Empty(result.OptimizedImages);
    }

    // ──────────────────────────────────────────────────────────
    //  2. Small image — skips optimization when already optimal
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void OptimizeImages_SmallImage_SkipsOptimization()
    {
        // Create a small PNG image (100x100) displayed at larger size (2x2 inches)
        var path = CreatePptxWithImage(
            width: 100,
            height: 100,
            format: MagickFormat.Png,
            displayWidthEmu: Emu.Inches2,
            displayHeightEmu: Emu.Inches2);

        var result = Service.OptimizeImages(path, targetDpi: 150);

        Assert.True(result.Success);
        // Image is already smaller than display size at 150 DPI, so should be skipped
        Assert.Equal(0, result.ImagesProcessed);
        Assert.True(result.ImagesSkipped > 0);
        Assert.Single(result.OptimizedImages);
        Assert.Equal(0, result.OptimizedImages[0].BytesSaved);
        Assert.Contains("skipped", result.OptimizedImages[0].Action, StringComparison.OrdinalIgnoreCase);
    }

    // ──────────────────────────────────────────────────────────
    //  3. JPEG recompression — verifies quality reduction
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void OptimizeImages_JpegRecompression_ReducesFileSize()
    {
        // Create a large JPEG at 100% quality with enough content to see compression
        var path = CreatePptxWithImage(
            width: 2000,
            height: 1500,
            format: MagickFormat.Jpeg,
            jpegQuality: 100,
            displayWidthEmu: Emu.Inches4,
            displayHeightEmu: Emu.Inches3);

        var result = Service.OptimizeImages(path, targetDpi: 150, jpegQuality: 85);

        Assert.True(result.Success);
        Assert.True(result.ImagesProcessed > 0);
        Assert.True(result.TotalBytesSaved > 0);
        Assert.Single(result.OptimizedImages);

        var optimizedImage = result.OptimizedImages[0];
        Assert.True(optimizedImage.BytesSaved > 0);
        Assert.Contains("recompressed", optimizedImage.Action);
        Assert.Equal("Jpeg", optimizedImage.OriginalFormat);
        Assert.Equal("Jpeg", optimizedImage.OptimizedFormat);
        Assert.True(optimizedImage.OptimizedSizeBytes < optimizedImage.OriginalSizeBytes);
    }

    // ──────────────────────────────────────────────────────────
    //  4. File not found — proper error handling
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void OptimizeImages_FileNotFound_ReturnsError()
    {
        var nonExistentPath = Path.Join(Path.GetTempPath(), "nonexistent-" + Guid.NewGuid() + ".pptx");

        var exception = Assert.Throws<FileNotFoundException>(() => Service.OptimizeImages(nonExistentPath));
        Assert.Contains(nonExistentPath, exception.Message);
    }

    // ──────────────────────────────────────────────────────────
    //  5. Result structure — all fields populated correctly
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void OptimizeImages_ResultStructure_AllFieldsPopulated()
    {
        var path = CreatePptxWithImage(
            width: 1500,
            height: 1000,
            format: MagickFormat.Jpeg,
            jpegQuality: 100,
            displayWidthEmu: Emu.Inches3,
            displayHeightEmu: Emu.Inches2);

        var result = Service.OptimizeImages(path);

        // Verify all top-level fields are present
        Assert.True(result.Success);
        Assert.NotNull(result.FilePath);
        Assert.Equal(path, result.FilePath);
        Assert.True(result.ImagesProcessed >= 0);
        Assert.True(result.ImagesSkipped >= 0);
        Assert.True(result.TotalBytesBefore >= 0);
        Assert.True(result.TotalBytesAfter >= 0);
        Assert.True(result.TotalBytesSaved >= 0);
        Assert.NotNull(result.OptimizedImages);
        Assert.NotEmpty(result.OptimizedImages);
        Assert.NotNull(result.Validation);
        Assert.NotNull(result.Message);
        Assert.False(string.IsNullOrWhiteSpace(result.Message));

        // Verify OptimizedImageInfo fields
        var imageInfo = result.OptimizedImages[0];
        Assert.NotNull(imageInfo.ImagePath);
        Assert.NotNull(imageInfo.OriginalFormat);
        Assert.NotNull(imageInfo.OptimizedFormat);
        Assert.True(imageInfo.OriginalWidth > 0);
        Assert.True(imageInfo.OriginalHeight > 0);
        Assert.True(imageInfo.OptimizedWidth > 0);
        Assert.True(imageInfo.OptimizedHeight > 0);
        Assert.True(imageInfo.OriginalSizeBytes > 0);
        Assert.True(imageInfo.OptimizedSizeBytes > 0);
        Assert.NotNull(imageInfo.Action);

        // Verify validation status
        Assert.True(result.Validation.ErrorsBefore >= 0);
        Assert.True(result.Validation.ErrorsAfter >= 0);
    }

    // ──────────────────────────────────────────────────────────
    //  6. Custom parameters — targetDpi, jpegQuality work
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void OptimizeImages_CustomParameters_AppliedCorrectly()
    {
        // Create large image that will be downscaled
        var path = CreatePptxWithImage(
            width: 3000,
            height: 2000,
            format: MagickFormat.Jpeg,
            jpegQuality: 100,
            displayWidthEmu: Emu.Inches3,  // 3 inches
            displayHeightEmu: Emu.Inches2); // 2 inches

        // At 300 DPI, 3 inches = 900 pixels, 2 inches = 600 pixels
        // Image is 3000x2000, so should be downscaled
        var result = Service.OptimizeImages(path, targetDpi: 300, jpegQuality: 90);

        Assert.True(result.Success);
        Assert.True(result.ImagesProcessed > 0);
        
        var optimizedImage = result.OptimizedImages[0];
        Assert.True(optimizedImage.OptimizedWidth < optimizedImage.OriginalWidth);
        Assert.True(optimizedImage.OptimizedHeight < optimizedImage.OriginalHeight);
        Assert.Contains("downscaled", optimizedImage.Action);
    }

    // ──────────────────────────────────────────────────────────
    //  7. Format conversion — BMP to PNG
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void OptimizeImages_BmpToPng_ConvertsFormat()
    {
        var path = CreatePptxWithImage(
            width: 800,
            height: 600,
            format: MagickFormat.Bmp,
            displayWidthEmu: Emu.Inches3,
            displayHeightEmu: Emu.Inches2);

        var result = Service.OptimizeImages(path, convertFormats: true);

        Assert.True(result.Success);
        Assert.True(result.ImagesProcessed > 0);
        
        var optimizedImage = result.OptimizedImages[0];
        Assert.Equal("Bmp", optimizedImage.OriginalFormat);
        Assert.Equal("Png", optimizedImage.OptimizedFormat);
        Assert.Contains("converted", optimizedImage.Action);
        Assert.True(optimizedImage.BytesSaved > 0);
    }

    // ──────────────────────────────────────────────────────────
    //  8. Downscaling large image — reduces dimensions
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void OptimizeImages_LargeImage_Downscales()
    {
        // Create very large image (4000x3000) displayed at 2x1.5 inches
        var path = CreatePptxWithImage(
            width: 4000,
            height: 3000,
            format: MagickFormat.Png,
            displayWidthEmu: Emu.Inches2,
            displayHeightEmu: Emu.Inches1_5);

        var result = Service.OptimizeImages(path, targetDpi: 150);

        Assert.True(result.Success);
        Assert.True(result.ImagesProcessed > 0);
        
        var optimizedImage = result.OptimizedImages[0];
        // At 150 DPI: 2 inches = 300px, 1.5 inches = 225px
        // Image should be downscaled from 4000x3000 to approximately 300x225
        Assert.True(optimizedImage.OptimizedWidth < optimizedImage.OriginalWidth);
        Assert.True(optimizedImage.OptimizedHeight < optimizedImage.OriginalHeight);
        Assert.True(optimizedImage.OptimizedWidth <= 350); // Allow some margin
        Assert.True(optimizedImage.OptimizedHeight <= 275);
        Assert.Contains("downscaled", optimizedImage.Action);
        Assert.True(optimizedImage.BytesSaved > 0);
    }

    // ──────────────────────────────────────────────────────────
    //  9. ConvertFormats disabled — BMP stays BMP
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void OptimizeImages_ConvertFormatsDisabled_SkipsConversion()
    {
        var path = CreatePptxWithImage(
            width: 500,
            height: 400,
            format: MagickFormat.Bmp,
            displayWidthEmu: Emu.Inches2,
            displayHeightEmu: Emu.Inches1_5);

        var result = Service.OptimizeImages(path, convertFormats: false);

        Assert.True(result.Success);
        // BMP without downscaling or conversion should be skipped
        var optimizedImage = result.OptimizedImages[0];
        Assert.Equal("Bmp", optimizedImage.OriginalFormat);
        Assert.Equal("Bmp", optimizedImage.OptimizedFormat);
    }

    // ──────────────────────────────────────────────────────────
    //  10. Multiple images — processes all
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void OptimizeImages_MultipleImages_ProcessesAll()
    {
        var path = CreatePptxWithMultipleImages();

        var result = Service.OptimizeImages(path);

        Assert.True(result.Success);
        // Should process 3 images total
        Assert.Equal(3, result.OptimizedImages.Count);
        Assert.True(result.ImagesProcessed + result.ImagesSkipped == 3);
        Assert.True(result.TotalBytesBefore > 0);
    }

    // ──────────────────────────────────────────────────────────
    //  Helpers — create PPTX fixtures with images
    // ──────────────────────────────────────────────────────────

    /// <summary>
    /// Creates a PPTX with a single slide containing one image with specified properties.
    /// </summary>
    private string CreatePptxWithImage(
        int width,
        int height,
        MagickFormat format,
        long displayWidthEmu,
        long displayHeightEmu,
        int jpegQuality = 100)
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var (slideMasterPart, slideLayoutPart) = CreateMinimalMasterAndLayout(presentationPart);

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.AddPart(slideLayoutPart);

        // Generate test image with ImageMagick
        var imageBytes = CreateTestImageBytes(width, height, format, jpegQuality);
        
        // Add image part based on format
        ImagePart imagePart = format switch
        {
            MagickFormat.Jpeg => slidePart.AddImagePart(ImagePartType.Jpeg),
            MagickFormat.Png => slidePart.AddImagePart(ImagePartType.Png),
            MagickFormat.Bmp => slidePart.AddImagePart(ImagePartType.Bmp),
            MagickFormat.Tiff or MagickFormat.Tiff64 => slidePart.AddImagePart(ImagePartType.Tiff),
            _ => slidePart.AddImagePart(ImagePartType.Png)
        };
        using (var ms = new MemoryStream(imageBytes))
            imagePart.FeedData(ms);

        var relId = slidePart.GetIdOfPart(imagePart);
        slidePart.Slide = CreateSlideWithPicture(relId, displayWidthEmu, displayHeightEmu);

        var slideIdList = new SlideIdList(
            new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });

        FinalizePresentationPart(presentationPart, slideIdList, slideMasterPart);
        return path;
    }

    /// <summary>
    /// Creates a PPTX with 3 slides, each with a different image.
    /// </summary>
    private string CreatePptxWithMultipleImages()
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var (slideMasterPart, slideLayoutPart) = CreateMinimalMasterAndLayout(presentationPart);

        var slideIdList = new SlideIdList();
        uint slideId = 256;

        // Image 1: Large JPEG
        var slidePart1 = presentationPart.AddNewPart<SlidePart>();
        slidePart1.AddPart(slideLayoutPart);
        var image1Bytes = CreateTestImageBytes(2000, 1500, MagickFormat.Jpeg, 100);
        var imagePart1 = slidePart1.AddImagePart(ImagePartType.Jpeg);
        using (var ms = new MemoryStream(image1Bytes))
            imagePart1.FeedData(ms);
        slidePart1.Slide = CreateSlideWithPicture(slidePart1.GetIdOfPart(imagePart1), Emu.Inches3, Emu.Inches2);
        slideIdList.Append(new SlideId { Id = slideId++, RelationshipId = presentationPart.GetIdOfPart(slidePart1) });

        // Image 2: BMP to convert
        var slidePart2 = presentationPart.AddNewPart<SlidePart>();
        slidePart2.AddPart(slideLayoutPart);
        var image2Bytes = CreateTestImageBytes(800, 600, MagickFormat.Bmp, 100);
        var imagePart2 = slidePart2.AddImagePart(ImagePartType.Bmp);
        using (var ms = new MemoryStream(image2Bytes))
            imagePart2.FeedData(ms);
        slidePart2.Slide = CreateSlideWithPicture(slidePart2.GetIdOfPart(imagePart2), Emu.Inches2, Emu.Inches1_5);
        slideIdList.Append(new SlideId { Id = slideId++, RelationshipId = presentationPart.GetIdOfPart(slidePart2) });

        // Image 3: Small PNG (should be skipped)
        var slidePart3 = presentationPart.AddNewPart<SlidePart>();
        slidePart3.AddPart(slideLayoutPart);
        var image3Bytes = CreateTestImageBytes(100, 100, MagickFormat.Png, 100);
        var imagePart3 = slidePart3.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(image3Bytes))
            imagePart3.FeedData(ms);
        slidePart3.Slide = CreateSlideWithPicture(slidePart3.GetIdOfPart(imagePart3), Emu.Inches2, Emu.Inches2);
        slideIdList.Append(new SlideId { Id = slideId++, RelationshipId = presentationPart.GetIdOfPart(slidePart3) });

        FinalizePresentationPart(presentationPart, slideIdList, slideMasterPart);
        return path;
    }

    /// <summary>
    /// Creates a test image using ImageMagick with specified properties.
    /// Returns a byte array instead of a stream to avoid disposal issues.
    /// </summary>
    private static byte[] CreateTestImageBytes(int width, int height, MagickFormat format, int jpegQuality)
    {
        using var image = new MagickImage(MagickColors.Blue, (uint)width, (uint)height);
        
        // Add some variation to make the image compressible
        using var overlay = new MagickImage(MagickColors.Yellow, (uint)width / 4, (uint)height / 4);
        image.Composite(overlay, (int)(width * 0.25), (int)(height * 0.25), CompositeOperator.Over);
        
        image.Format = format;
        if (format == MagickFormat.Jpeg)
        {
            image.Quality = (uint)jpegQuality;
        }
        
        return image.ToByteArray();
    }
    
    /// <summary>
    /// Creates a test image using ImageMagick with specified properties.
    /// </summary>
    private static MemoryStream CreateTestImage(int width, int height, MagickFormat format, int jpegQuality)
    {
        var bytes = CreateTestImageBytes(width, height, format, jpegQuality);
        return new MemoryStream(bytes);
    }

    // ── Shared fixture helpers ──────────────────────────────

    private static (SlideMasterPart Master, SlideLayoutPart Layout) CreateMinimalMasterAndLayout(
        PresentationPart presentationPart)
    {
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

        slideLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()))
        { Type = SlideLayoutValues.Title };
        slideLayoutPart.SlideLayout.CommonSlideData!.Name = "Title Slide";
        slideLayoutPart.AddPart(slideMasterPart);

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new SlideLayoutIdList(
                new SlideLayoutId
                {
                    Id = 2049,
                    RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
                }));

        return (slideMasterPart, slideLayoutPart);
    }

    private static void FinalizePresentationPart(PresentationPart presentationPart,
        SlideIdList slideIdList, SlideMasterPart slideMasterPart)
    {
        var slideMasterIdList = new SlideMasterIdList(
            new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

        presentationPart.Presentation = new Presentation(
            slideIdList,
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });

        presentationPart.Presentation.InsertAt(slideMasterIdList, 0);
        presentationPart.Presentation.Save();
    }

    private static Slide CreateSlideWithPicture(string imageRelId, long widthEmu, long heightEmu)
    {
        return new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()),
                    new P.Picture(
                        new P.NonVisualPictureProperties(
                            new P.NonVisualDrawingProperties { Id = 2, Name = "Image 1" },
                            new P.NonVisualPictureDrawingProperties(
                                new A.PictureLocks { NoChangeAspect = true }),
                            new ApplicationNonVisualDrawingProperties()),
                        new P.BlipFill(
                            new A.Blip { Embed = imageRelId },
                            new A.Stretch(new A.FillRectangle())),
                        new P.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = Emu.OneInch, Y = Emu.OneInch },
                                new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                            new A.PresetGeometry(new A.AdjustValueList())
                            { Preset = A.ShapeTypeValues.Rectangle })))));
    }
}
