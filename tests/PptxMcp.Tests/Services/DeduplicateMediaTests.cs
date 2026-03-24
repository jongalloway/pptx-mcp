using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxMcp.Tests.Services;

/// <summary>
/// Service-level tests for DeduplicateMedia (Issue #84 — Deduplicate identical media).
/// Validates deduplication correctness, reference integrity, validation, and round-trip safety.
/// </summary>
[Trait("Category", "Unit")]
public class DeduplicateMediaTests : PptxTestBase
{
    // ──────────────────────────────────────────────────────────
    //  1. No duplicates — file unchanged, 0 parts removed
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void DeduplicateMedia_NoDuplicates_ReturnsNoOp()
    {
        var path = CreatePptxWithUniqueImages(imageCount: 2);

        var result = Service.DeduplicateMedia(path);

        Assert.True(result.Success);
        Assert.Equal(0, result.DuplicateGroupsFound);
        Assert.Equal(0, result.PartsRemoved);
        Assert.Equal(0, result.BytesSaved);
        Assert.Empty(result.Groups);
    }

    // ──────────────────────────────────────────────────────────
    //  2. Two identical images on different slides — one removed
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void DeduplicateMedia_TwoIdenticalImages_RemovesOne()
    {
        var path = CreatePptxWithDuplicateImage();

        var result = Service.DeduplicateMedia(path);

        Assert.True(result.Success);
        Assert.Equal(1, result.DuplicateGroupsFound);
        Assert.Equal(1, result.PartsRemoved);
        Assert.True(result.BytesSaved > 0);
        Assert.Single(result.Groups);

        var group = result.Groups[0];
        Assert.Single(group.RemovedPartUris);
        Assert.NotEqual(group.CanonicalPartUri, group.RemovedPartUris[0]);
    }

    // ──────────────────────────────────────────────────────────
    //  3. Multiple duplicate groups — all deduplicated
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void DeduplicateMedia_MultipleDuplicateGroups_AllDeduplicated()
    {
        var path = CreatePptxWithMultipleDuplicateGroups();

        var result = Service.DeduplicateMedia(path);

        Assert.True(result.Success);
        Assert.Equal(2, result.DuplicateGroupsFound);
        Assert.Equal(2, result.PartsRemoved);
        Assert.True(result.BytesSaved > 0);
        Assert.Equal(2, result.Groups.Count);
    }

    // ──────────────────────────────────────────────────────────
    //  4. Validates before and after — no new errors
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void DeduplicateMedia_ValidatesBeforeAndAfter()
    {
        var path = CreatePptxWithDuplicateImage();

        var result = Service.DeduplicateMedia(path);

        Assert.True(result.Success);
        Assert.NotNull(result.Validation);
        Assert.True(result.Validation.ErrorsBefore >= 0);
        Assert.True(result.Validation.ErrorsAfter <= result.Validation.ErrorsBefore,
            "Deduplication should not introduce new validation errors.");
    }

    // ──────────────────────────────────────────────────────────
    //  5. Round-trip — file opens in OpenXml after dedup
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void DeduplicateMedia_FileOpensInOpenXml_AfterDedup()
    {
        var path = CreatePptxWithDuplicateImage();

        var result = Service.DeduplicateMedia(path);
        Assert.True(result.Success);

        // Re-open and verify structural integrity
        using var doc = PresentationDocument.Open(path, false);
        var presentationPart = doc.PresentationPart;
        Assert.NotNull(presentationPart);

        var slideIdList = presentationPart.Presentation.SlideIdList;
        Assert.NotNull(slideIdList);
        Assert.NotEmpty(slideIdList.Elements<SlideId>());
    }

    // ──────────────────────────────────────────────────────────
    //  6. References updated — Blip.Embed points to canonical
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void DeduplicateMedia_BlipReferencesPointToCanonical()
    {
        var path = CreatePptxWithDuplicateImage();

        var result = Service.DeduplicateMedia(path);
        Assert.True(result.Success);
        Assert.Single(result.Groups);

        var canonicalUri = result.Groups[0].CanonicalPartUri;

        // Verify all slide images reference the canonical part
        using var doc = PresentationDocument.Open(path, false);
        var presentationPart = doc.PresentationPart!;

        foreach (var slideId in presentationPart.Presentation.SlideIdList!.Elements<SlideId>())
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!);
            var blips = slidePart.Slide.Descendants<Blip>();
            foreach (var blip in blips)
            {
                if (blip.Embed?.Value is { } relId &&
                    slidePart.TryGetPartById(relId, out var part) &&
                    part is ImagePart imagePart)
                {
                    Assert.Equal(canonicalUri, imagePart.Uri.ToString());
                }
            }
        }
    }

    // ──────────────────────────────────────────────────────────
    //  7. Image on layout — still handled
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void DeduplicateMedia_ImageOnLayout_StillDeduplicated()
    {
        var path = CreatePptxWithDuplicateImageOnLayout();

        var result = Service.DeduplicateMedia(path);

        Assert.True(result.Success);
        Assert.Equal(1, result.DuplicateGroupsFound);
        Assert.Equal(1, result.PartsRemoved);
    }

    // ──────────────────────────────────────────────────────────
    //  8. FilePath matches input
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void DeduplicateMedia_FilePath_MatchesInput()
    {
        var path = CreatePptxWithDuplicateImage();

        var result = Service.DeduplicateMedia(path);

        Assert.Equal(path, result.FilePath);
    }

    // ──────────────────────────────────────────────────────────
    //  9. Message is populated
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void DeduplicateMedia_Message_IsNotEmpty()
    {
        var path = CreatePptxWithDuplicateImage();

        var result = Service.DeduplicateMedia(path);

        Assert.False(string.IsNullOrWhiteSpace(result.Message));
    }

    // ──────────────────────────────────────────────────────────
    //  10. Media count decreases after dedup
    // ──────────────────────────────────────────────────────────

    [Fact]
    public void DeduplicateMedia_MediaCountDecreases()
    {
        var path = CreatePptxWithDuplicateImage();

        var analysisBefore = Service.AnalyzeMedia(path);
        var result = Service.DeduplicateMedia(path);
        var analysisAfter = Service.AnalyzeMedia(path);

        Assert.True(result.Success);
        Assert.True(analysisAfter.TotalMediaCount < analysisBefore.TotalMediaCount,
            "Media count should decrease after deduplication.");
    }

    // ──────────────────────────────────────────────────────────
    //  Helpers — create PPTX fixtures with specific media patterns
    // ──────────────────────────────────────────────────────────

    /// <summary>
    /// Creates a PPTX with the specified number of slides, each having a unique image.
    /// </summary>
    private string CreatePptxWithUniqueImages(int imageCount)
    {
        var path = System.IO.Path.Join(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var (slideMasterPart, slideLayoutPart) = CreateMinimalMasterAndLayout(presentationPart);

        var slideIdList = new SlideIdList();
        uint slideId = 256;

        for (int i = 0; i < imageCount; i++)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(slideLayoutPart);

            // Create a unique image (different pixel for each)
            var imagePart = slidePart.AddImagePart(ImagePartType.Png);
            var imageBytes = CreatePngImage((byte)(i + 1));
            using (var ms = new MemoryStream(imageBytes))
                imagePart.FeedData(ms);

            var relId = slidePart.GetIdOfPart(imagePart);
            slidePart.Slide = CreateSlideWithPicture(relId);

            slideIdList.Append(new SlideId
            {
                Id = slideId++,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        FinalizePresentationPart(presentationPart, slideIdList, slideMasterPart);
        return path;
    }

    /// <summary>
    /// Creates a PPTX with 2 slides, both referencing identical (but separate) image parts.
    /// </summary>
    private string CreatePptxWithDuplicateImage()
    {
        var path = System.IO.Path.Join(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        var imageBytes = CreatePngImage(0xAA);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var (slideMasterPart, slideLayoutPart) = CreateMinimalMasterAndLayout(presentationPart);

        var slideIdList = new SlideIdList();
        uint slideId = 256;

        for (int i = 0; i < 2; i++)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(slideLayoutPart);

            var imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using (var ms = new MemoryStream(imageBytes))
                imagePart.FeedData(ms);

            var relId = slidePart.GetIdOfPart(imagePart);
            slidePart.Slide = CreateSlideWithPicture(relId);

            slideIdList.Append(new SlideId
            {
                Id = slideId++,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        FinalizePresentationPart(presentationPart, slideIdList, slideMasterPart);
        return path;
    }

    /// <summary>
    /// Creates a PPTX with 4 slides and 2 duplicate groups (2 pairs of identical images).
    /// </summary>
    private string CreatePptxWithMultipleDuplicateGroups()
    {
        var path = System.IO.Path.Join(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        var imageA = CreatePngImage(0xAA);
        var imageB = CreatePngImage(0xBB);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var (slideMasterPart, slideLayoutPart) = CreateMinimalMasterAndLayout(presentationPart);

        var slideIdList = new SlideIdList();
        uint slideId = 256;

        // Group A: slides 1 and 2 use imageA
        // Group B: slides 3 and 4 use imageB
        var images = new[] { imageA, imageA, imageB, imageB };
        for (int i = 0; i < 4; i++)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(slideLayoutPart);

            var imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using (var ms = new MemoryStream(images[i]))
                imagePart.FeedData(ms);

            var relId = slidePart.GetIdOfPart(imagePart);
            slidePart.Slide = CreateSlideWithPicture(relId);

            slideIdList.Append(new SlideId
            {
                Id = slideId++,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        FinalizePresentationPart(presentationPart, slideIdList, slideMasterPart);
        return path;
    }

    /// <summary>
    /// Creates a PPTX where slide and layout both reference identical images.
    /// </summary>
    private string CreatePptxWithDuplicateImageOnLayout()
    {
        var path = System.IO.Path.Join(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        var imageBytes = CreatePngImage(0xCC);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var (slideMasterPart, slideLayoutPart) = CreateMinimalMasterAndLayout(presentationPart);

        // Add image to layout
        var layoutImagePart = slideLayoutPart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(imageBytes))
            layoutImagePart.FeedData(ms);

        var layoutRelId = slideLayoutPart.GetIdOfPart(layoutImagePart);
        AddPictureToShapeTree(slideLayoutPart.SlideLayout.CommonSlideData!.ShapeTree!, layoutRelId, 2);

        // Add identical image to slide
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.AddPart(slideLayoutPart);

        var slideImagePart = slidePart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(imageBytes))
            slideImagePart.FeedData(ms);

        var slideRelId = slidePart.GetIdOfPart(slideImagePart);
        slidePart.Slide = CreateSlideWithPicture(slideRelId);

        var slideIdList = new SlideIdList(
            new SlideId
            {
                Id = 256,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });

        FinalizePresentationPart(presentationPart, slideIdList, slideMasterPart);
        return path;
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

    private static Slide CreateSlideWithPicture(string imageRelId)
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
                            new Blip { Embed = imageRelId },
                            new Stretch(new FillRectangle())),
                        new P.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0, Y = 0 },
                                new A.Extents { Cx = 1000000, Cy = 1000000 }),
                            new A.PresetGeometry(new A.AdjustValueList())
                            { Preset = A.ShapeTypeValues.Rectangle })))));
    }

    private static void AddPictureToShapeTree(ShapeTree shapeTree, string imageRelId, uint shapeId)
    {
        shapeTree.Append(
            new P.Picture(
                new P.NonVisualPictureProperties(
                    new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Image {shapeId}" },
                    new P.NonVisualPictureDrawingProperties(
                        new A.PictureLocks { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new P.BlipFill(
                    new Blip { Embed = imageRelId },
                    new Stretch(new FillRectangle())),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 0, Y = 0 },
                        new A.Extents { Cx = 1000000, Cy = 1000000 }),
                    new A.PresetGeometry(new A.AdjustValueList())
                    { Preset = A.ShapeTypeValues.Rectangle })));
    }

    /// <summary>
    /// Creates a minimal valid PNG image with a single pixel of the specified color value.
    /// </summary>
    private static byte[] CreatePngImage(byte colorValue)
    {
        // Minimal 1x1 PNG with varying pixel data to produce different hashes.
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms);

        // PNG signature
        writer.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });

        // IHDR chunk: 1x1, 8-bit RGB
        WriteChunk(writer, "IHDR", new byte[]
        {
            0, 0, 0, 1, // width
            0, 0, 0, 1, // height
            8,           // bit depth
            2,           // color type (RGB)
            0, 0, 0      // compression, filter, interlace
        });

        // IDAT chunk: deflated scanline (filter byte + 3 RGB bytes)
        var rawScanline = new byte[] { 0, colorValue, colorValue, colorValue };
        var deflated = DeflateCompress(rawScanline);
        WriteChunk(writer, "IDAT", deflated);

        // IEND chunk
        WriteChunk(writer, "IEND", []);

        return ms.ToArray();
    }

    private static void WriteChunk(BinaryWriter writer, string type, byte[] data)
    {
        var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        writer.Write(ToBigEndian(data.Length));
        writer.Write(typeBytes);
        writer.Write(data);

        // CRC32 over type + data
        var crcData = new byte[typeBytes.Length + data.Length];
        typeBytes.CopyTo(crcData, 0);
        data.CopyTo(crcData, typeBytes.Length);
        writer.Write(ToBigEndian((int)Crc32(crcData)));
    }

    private static byte[] ToBigEndian(int value) =>
        [(byte)(value >> 24), (byte)(value >> 16), (byte)(value >> 8), (byte)value];

    private static byte[] DeflateCompress(byte[] data)
    {
        using var ms = new MemoryStream();
        // zlib header
        ms.WriteByte(0x78);
        ms.WriteByte(0x01);
        using (var deflate = new System.IO.Compression.DeflateStream(ms,
            System.IO.Compression.CompressionLevel.NoCompression, leaveOpen: true))
        {
            deflate.Write(data, 0, data.Length);
        }
        // Adler32 checksum
        var adler = Adler32(data);
        ms.Write(ToBigEndian((int)adler));
        return ms.ToArray();
    }

    private static uint Adler32(byte[] data)
    {
        uint a = 1, b = 0;
        foreach (var d in data)
        {
            a = (a + d) % 65521;
            b = (b + a) % 65521;
        }
        return (b << 16) | a;
    }

    private static uint Crc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (var b in data)
        {
            crc ^= b;
            for (int i = 0; i < 8; i++)
                crc = (crc >> 1) ^ (0xEDB88320 & ~((crc & 1) - 1));
        }
        return crc ^ 0xFFFFFFFF;
    }
}
