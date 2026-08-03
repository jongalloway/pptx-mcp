using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxTools.Services;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxTools.Tests.Services;

[Trait("Category", "Unit")]
public class VideoMetadataTests : PptxTestBase
{
    // ── Minimal MP4 generation ─────────────────────────────────────────
    // A valid minimal MP4 needs: ftyp box + moov box with mvhd + trak (tkhd + mdia (mdhd + hdlr + minf (stbl (stsd))))

    /// <summary>
    /// Build a minimal valid MP4 byte sequence with a single H.264 video track.
    /// This is a hand-crafted ISOBMFF structure with the minimum boxes needed
    /// for SharpMP4 to parse track metadata.
    /// </summary>
    private static byte[] BuildMinimalMp4(
        ushort width = 1920,
        ushort height = 1080,
        uint timescale = 90000,
        uint duration = 450000, // 5 seconds at 90kHz
        uint movieTimescale = 1000,
        uint movieDuration = 5000)
    {
        using var ms = new MemoryStream();
        using var bw = new BinaryWriter(ms);

        // ftyp box
        WriteFtypBox(bw);

        // moov box (movie container)
        var moovContent = new MemoryStream();
        using (var moovWriter = new BinaryWriter(moovContent, System.Text.Encoding.UTF8, leaveOpen: true))
        {
            // mvhd (Movie Header)
            WriteMvhdBox(moovWriter, movieTimescale, movieDuration);

            // trak (Track container)
            var trakContent = new MemoryStream();
            using (var trakWriter = new BinaryWriter(trakContent, System.Text.Encoding.UTF8, leaveOpen: true))
            {
                // tkhd (Track Header)
                WriteTkhdBox(trakWriter, width, height, movieDuration);

                // mdia (Media container)
                var mdiaContent = new MemoryStream();
                using (var mdiaWriter = new BinaryWriter(mdiaContent, System.Text.Encoding.UTF8, leaveOpen: true))
                {
                    // mdhd (Media Header)
                    WriteMdhdBox(mdiaWriter, timescale, duration);

                    // hdlr (Handler Reference - video)
                    WriteHdlrBox(mdiaWriter, "vide", "VideoHandler");

                    // minf (Media Information)
                    var minfContent = new MemoryStream();
                    using (var minfWriter = new BinaryWriter(minfContent, System.Text.Encoding.UTF8, leaveOpen: true))
                    {
                        // vmhd (Video Media Header)
                        WriteVmhdBox(minfWriter);

                        // stbl (Sample Table)
                        var stblContent = new MemoryStream();
                        using (var stblWriter = new BinaryWriter(stblContent, System.Text.Encoding.UTF8, leaveOpen: true))
                        {
                            // stsd with avc1 visual sample entry
                            WriteStsdVideoBox(stblWriter, width, height);

                            // stts (empty but required)
                            WriteEmptySttsBox(stblWriter);

                            // stsc (empty but required)
                            WriteEmptyStscBox(stblWriter);

                            // stsz (empty)
                            WriteEmptyStszBox(stblWriter);

                            // stco (empty)
                            WriteEmptyStcoBox(stblWriter);
                        }
                        WriteContainerBox(minfWriter, "stbl", stblContent.ToArray());
                    }
                    WriteContainerBox(mdiaWriter, "minf", minfContent.ToArray());
                }
                WriteContainerBox(trakWriter, "mdia", mdiaContent.ToArray());
            }
            WriteContainerBox(moovWriter, "trak", trakContent.ToArray());
        }
        WriteContainerBox(bw, "moov", moovContent.ToArray());

        return ms.ToArray();
    }

    /// <summary>Build a minimal MP4 with a single AAC audio track.</summary>
    private static byte[] BuildMinimalAudioMp4(
        ushort channelCount = 2,
        uint sampleRate = 44100,
        uint timescale = 44100,
        uint duration = 220500, // 5 seconds at 44100Hz
        uint movieTimescale = 1000,
        uint movieDuration = 5000)
    {
        using var ms = new MemoryStream();
        using var bw = new BinaryWriter(ms);

        WriteFtypBox(bw);

        var moovContent = new MemoryStream();
        using (var moovWriter = new BinaryWriter(moovContent, System.Text.Encoding.UTF8, leaveOpen: true))
        {
            WriteMvhdBox(moovWriter, movieTimescale, movieDuration);

            var trakContent = new MemoryStream();
            using (var trakWriter = new BinaryWriter(trakContent, System.Text.Encoding.UTF8, leaveOpen: true))
            {
                WriteTkhdBox(trakWriter, 0, 0, movieDuration);

                var mdiaContent = new MemoryStream();
                using (var mdiaWriter = new BinaryWriter(mdiaContent, System.Text.Encoding.UTF8, leaveOpen: true))
                {
                    WriteMdhdBox(mdiaWriter, timescale, duration);
                    WriteHdlrBox(mdiaWriter, "soun", "SoundHandler");

                    var minfContent = new MemoryStream();
                    using (var minfWriter = new BinaryWriter(minfContent, System.Text.Encoding.UTF8, leaveOpen: true))
                    {
                        WriteSmhdBox(minfWriter);

                        var stblContent = new MemoryStream();
                        using (var stblWriter = new BinaryWriter(stblContent, System.Text.Encoding.UTF8, leaveOpen: true))
                        {
                            WriteStsdAudioBox(stblWriter, channelCount, sampleRate);
                            WriteEmptySttsBox(stblWriter);
                            WriteEmptyStscBox(stblWriter);
                            WriteEmptyStszBox(stblWriter);
                            WriteEmptyStcoBox(stblWriter);
                        }
                        WriteContainerBox(minfWriter, "stbl", stblContent.ToArray());
                    }
                    WriteContainerBox(mdiaWriter, "minf", minfContent.ToArray());
                }
                WriteContainerBox(trakWriter, "mdia", mdiaContent.ToArray());
            }
            WriteContainerBox(moovWriter, "trak", trakContent.ToArray());
        }
        WriteContainerBox(bw, "moov", moovContent.ToArray());

        return ms.ToArray();
    }

    // ── Box writing helpers ────────────────────────────────────────────

    private static void WriteBigEndian(BinaryWriter bw, uint value)
    {
        bw.Write((byte)((value >> 24) & 0xFF));
        bw.Write((byte)((value >> 16) & 0xFF));
        bw.Write((byte)((value >> 8) & 0xFF));
        bw.Write((byte)(value & 0xFF));
    }

    private static void WriteBigEndian(BinaryWriter bw, ushort value)
    {
        bw.Write((byte)((value >> 8) & 0xFF));
        bw.Write((byte)(value & 0xFF));
    }

    private static void WriteFourCC(BinaryWriter bw, string fourcc)
    {
        bw.Write(System.Text.Encoding.ASCII.GetBytes(fourcc));
    }

    private static void WriteContainerBox(BinaryWriter bw, string type, byte[] content)
    {
        WriteBigEndian(bw, (uint)(8 + content.Length));
        WriteFourCC(bw, type);
        bw.Write(content);
    }

    private static void WriteFtypBox(BinaryWriter bw)
    {
        // ftyp: size(4) + "ftyp"(4) + major_brand(4) + minor_version(4) + compatible_brands(4)
        WriteBigEndian(bw, 20u);
        WriteFourCC(bw, "ftyp");
        WriteFourCC(bw, "isom");
        WriteBigEndian(bw, 0u);      // minor version
        WriteFourCC(bw, "isom");     // compatible brand
    }

    private static void WriteMvhdBox(BinaryWriter bw, uint timescale, uint duration)
    {
        // mvhd version 0: 108 bytes total
        WriteBigEndian(bw, 108u);
        WriteFourCC(bw, "mvhd");
        WriteBigEndian(bw, 0u);          // version + flags
        WriteBigEndian(bw, 0u);          // creation_time
        WriteBigEndian(bw, 0u);          // modification_time
        WriteBigEndian(bw, timescale);
        WriteBigEndian(bw, duration);
        WriteBigEndian(bw, 0x00010000u); // rate (1.0)
        WriteBigEndian(bw, 0x01000000u); // volume (1.0) + reserved(2)
        bw.Write(new byte[8]);           // reserved
        // identity matrix (36 bytes)
        WriteBigEndian(bw, 0x00010000u); bw.Write(new byte[12]);
        WriteBigEndian(bw, 0x00010000u); bw.Write(new byte[12]);
        WriteBigEndian(bw, 0x40000000u);
        bw.Write(new byte[24]);          // pre_defined
        WriteBigEndian(bw, 2u);          // next_track_ID
    }

    private static void WriteTkhdBox(BinaryWriter bw, ushort width, ushort height, uint duration)
    {
        // tkhd version 0: 92 bytes total
        WriteBigEndian(bw, 92u);
        WriteFourCC(bw, "tkhd");
        WriteBigEndian(bw, 0x00000003u); // version 0 + flags (track_enabled | track_in_movie)
        WriteBigEndian(bw, 0u);          // creation_time
        WriteBigEndian(bw, 0u);          // modification_time
        WriteBigEndian(bw, 1u);          // track_ID
        WriteBigEndian(bw, 0u);          // reserved
        WriteBigEndian(bw, duration);
        bw.Write(new byte[8]);           // reserved
        WriteBigEndian(bw, 0u);          // layer(2) + alternate_group(2)
        WriteBigEndian(bw, 0u);          // volume(2) + reserved(2)
        // identity matrix (36 bytes)
        WriteBigEndian(bw, 0x00010000u); bw.Write(new byte[12]);
        WriteBigEndian(bw, 0x00010000u); bw.Write(new byte[12]);
        WriteBigEndian(bw, 0x40000000u);
        // width and height as 16.16 fixed point
        WriteBigEndian(bw, (uint)width << 16);
        WriteBigEndian(bw, (uint)height << 16);
    }

    private static void WriteMdhdBox(BinaryWriter bw, uint timescale, uint duration)
    {
        // mdhd version 0: 32 bytes total
        WriteBigEndian(bw, 32u);
        WriteFourCC(bw, "mdhd");
        WriteBigEndian(bw, 0u);          // version + flags
        WriteBigEndian(bw, 0u);          // creation_time
        WriteBigEndian(bw, 0u);          // modification_time
        WriteBigEndian(bw, timescale);
        WriteBigEndian(bw, duration);
        WriteBigEndian(bw, 0x55C40000u); // language ("und") + pre_defined
    }

    private static void WriteHdlrBox(BinaryWriter bw, string handlerType, string name)
    {
        byte[] nameBytes = System.Text.Encoding.UTF8.GetBytes(name + "\0");
        uint size = (uint)(32 + nameBytes.Length);
        WriteBigEndian(bw, size);
        WriteFourCC(bw, "hdlr");
        WriteBigEndian(bw, 0u);          // version + flags
        WriteBigEndian(bw, 0u);          // pre_defined
        WriteFourCC(bw, handlerType);
        bw.Write(new byte[12]);          // reserved
        bw.Write(nameBytes);
    }

    private static void WriteVmhdBox(BinaryWriter bw)
    {
        WriteBigEndian(bw, 20u);
        WriteFourCC(bw, "vmhd");
        WriteBigEndian(bw, 0x00000001u); // version 0 + flags=1
        bw.Write(new byte[8]);           // graphicsmode + opcolor
    }

    private static void WriteSmhdBox(BinaryWriter bw)
    {
        WriteBigEndian(bw, 16u);
        WriteFourCC(bw, "smhd");
        WriteBigEndian(bw, 0u);          // version + flags
        WriteBigEndian(bw, 0u);          // balance + reserved
    }

    private static void WriteStsdVideoBox(BinaryWriter bw, ushort width, ushort height)
    {
        // stsd with one avc1 VisualSampleEntry
        var entryContent = new MemoryStream();
        using (var ew = new BinaryWriter(entryContent, System.Text.Encoding.UTF8, leaveOpen: true))
        {
            // VisualSampleEntry (avc1):
            // 6 bytes reserved + 2 bytes data_ref_index
            ew.Write(new byte[6]);
            WriteBigEndian(ew, (ushort)1);  // data_reference_index
            // 2 pre_defined + 2 reserved + 12 pre_defined
            ew.Write(new byte[16]);
            WriteBigEndian(ew, width);
            WriteBigEndian(ew, height);
            WriteBigEndian(ew, 0x00480000u); // horizresolution (72 dpi)
            WriteBigEndian(ew, 0x00480000u); // vertresolution (72 dpi)
            WriteBigEndian(ew, 0u);          // reserved
            WriteBigEndian(ew, (ushort)1);   // frame_count
            ew.Write(new byte[32]);          // compressorname
            WriteBigEndian(ew, (ushort)0x0018); // depth
            WriteBigEndian(ew, unchecked((ushort)0xFFFF)); // pre_defined = -1
        }
        byte[] avc1Content = entryContent.ToArray();
        uint avc1Size = (uint)(8 + avc1Content.Length);

        // stsd: size + "stsd" + version/flags + entry_count + entries
        uint stsdSize = 16 + avc1Size;
        WriteBigEndian(bw, stsdSize);
        WriteFourCC(bw, "stsd");
        WriteBigEndian(bw, 0u);          // version + flags
        WriteBigEndian(bw, 1u);          // entry_count
        WriteBigEndian(bw, avc1Size);
        WriteFourCC(bw, "avc1");
        bw.Write(avc1Content);
    }

    private static void WriteStsdAudioBox(BinaryWriter bw, ushort channelCount, uint sampleRate)
    {
        var entryContent = new MemoryStream();
        using (var ew = new BinaryWriter(entryContent, System.Text.Encoding.UTF8, leaveOpen: true))
        {
            // AudioSampleEntry (mp4a):
            ew.Write(new byte[6]);           // reserved
            WriteBigEndian(ew, (ushort)1);   // data_reference_index
            ew.Write(new byte[8]);           // reserved
            WriteBigEndian(ew, channelCount);
            WriteBigEndian(ew, (ushort)16);  // sample_size
            WriteBigEndian(ew, (ushort)0);   // pre_defined
            WriteBigEndian(ew, (ushort)0);   // reserved
            WriteBigEndian(ew, sampleRate << 16); // sample_rate as 16.16 fixed point
        }
        byte[] mp4aContent = entryContent.ToArray();
        uint mp4aSize = (uint)(8 + mp4aContent.Length);

        uint stsdSize = 16 + mp4aSize;
        WriteBigEndian(bw, stsdSize);
        WriteFourCC(bw, "stsd");
        WriteBigEndian(bw, 0u);
        WriteBigEndian(bw, 1u);
        WriteBigEndian(bw, mp4aSize);
        WriteFourCC(bw, "mp4a");
        bw.Write(mp4aContent);
    }

    private static void WriteEmptySttsBox(BinaryWriter bw)
    {
        WriteBigEndian(bw, 16u);
        WriteFourCC(bw, "stts");
        WriteBigEndian(bw, 0u);
        WriteBigEndian(bw, 0u);
    }

    private static void WriteEmptyStscBox(BinaryWriter bw)
    {
        WriteBigEndian(bw, 16u);
        WriteFourCC(bw, "stsc");
        WriteBigEndian(bw, 0u);
        WriteBigEndian(bw, 0u);
    }

    private static void WriteEmptyStszBox(BinaryWriter bw)
    {
        WriteBigEndian(bw, 20u);
        WriteFourCC(bw, "stsz");
        WriteBigEndian(bw, 0u);
        WriteBigEndian(bw, 0u);
        WriteBigEndian(bw, 0u);
    }

    private static void WriteEmptyStcoBox(BinaryWriter bw)
    {
        WriteBigEndian(bw, 16u);
        WriteFourCC(bw, "stco");
        WriteBigEndian(bw, 0u);
        WriteBigEndian(bw, 0u);
    }

    // ── PPTX with embedded video helper ────────────────────────────────

    private string CreatePptxWithEmbeddedVideo(byte[]? videoBytes = null)
    {
        videoBytes ??= BuildMinimalMp4();
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        presentationPart.Presentation = new Presentation(
            new SlideMasterIdList(),
            new SlideIdList(),
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });

        var slidePart = presentationPart.AddNewPart<SlidePart>("rId2");
        slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));

        var slideIdList = presentationPart.Presentation.SlideIdList!;
        slideIdList.Append(new SlideId { Id = 256, RelationshipId = "rId2" });

        // Embed the video as a MediaDataPart
        var mediaDataPart = doc.CreateMediaDataPart("video/mp4", ".mp4");
        using (var mediaStream = mediaDataPart.GetStream(FileMode.Create))
        {
            mediaStream.Write(videoBytes, 0, videoBytes.Length);
        }

        // Create a reference from the slide to the media
        slidePart.AddMediaReferenceRelationship(mediaDataPart);

        presentationPart.Presentation.Save();
        return path;
    }

    private string CreatePptxWithEmbeddedAudio(byte[]? audioBytes = null)
    {
        audioBytes ??= BuildMinimalAudioMp4();
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        TrackTempFile(path);

        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        presentationPart.Presentation = new Presentation(
            new SlideMasterIdList(),
            new SlideIdList(),
            new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });

        var slidePart = presentationPart.AddNewPart<SlidePart>("rId2");
        slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));

        var slideIdList = presentationPart.Presentation.SlideIdList!;
        slideIdList.Append(new SlideId { Id = 256, RelationshipId = "rId2" });

        var mediaDataPart = doc.CreateMediaDataPart("audio/mp4", ".m4a");
        using (var mediaStream = mediaDataPart.GetStream(FileMode.Create))
        {
            mediaStream.Write(audioBytes, 0, audioBytes.Length);
        }

        slidePart.AddMediaReferenceRelationship(mediaDataPart);

        presentationPart.Presentation.Save();
        return path;
    }

    // ── Tests: PPTX with no video ──────────────────────────────────────

    [Fact]
    public void AnalyzeVideoMetadata_NoVideo_ReturnsSuccessWithZeroParts()
    {
        var path = CreateMinimalPptx();

        var result = Service.AnalyzeVideoMetadata(path);

        Assert.True(result.Success);
        Assert.Equal(0, result.VideoPartsFound);
        Assert.Equal(0, result.TotalTracks);
        Assert.Empty(result.Parts);
    }

    [Fact]
    public void AnalyzeVideoMetadata_NoVideo_ReturnsFilePath()
    {
        var path = CreateMinimalPptx();

        var result = Service.AnalyzeVideoMetadata(path);

        Assert.Equal(path, result.FilePath);
    }

    [Fact]
    public void AnalyzeVideoMetadata_NoVideo_MessageIndicatesNoMedia()
    {
        var path = CreateMinimalPptx();

        var result = Service.AnalyzeVideoMetadata(path);

        Assert.Contains("No video", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ── Tests: PPTX with embedded MP4 video ────────────────────────────

    [Fact]
    public void AnalyzeVideoMetadata_EmbeddedMp4_FindsOnePart()
    {
        var path = CreatePptxWithEmbeddedVideo();

        var result = Service.AnalyzeVideoMetadata(path);

        Assert.True(result.Success);
        Assert.Equal(1, result.VideoPartsFound);
        Assert.NotEmpty(result.Parts);
    }

    [Fact]
    public void AnalyzeVideoMetadata_EmbeddedMp4_ExtractsVideoTrack()
    {
        var path = CreatePptxWithEmbeddedVideo();

        var result = Service.AnalyzeVideoMetadata(path);

        var part = Assert.Single(result.Parts);
        Assert.Contains(part.Tracks, t => t.TrackType == "video");
    }

    [Fact]
    public void AnalyzeVideoMetadata_EmbeddedMp4_ExtractsCodecName()
    {
        var path = CreatePptxWithEmbeddedVideo();

        var result = Service.AnalyzeVideoMetadata(path);

        var videoTrack = result.Parts[0].Tracks.First(t => t.TrackType == "video");
        Assert.Equal("avc1", videoTrack.Codec);
    }

    [Fact]
    public void AnalyzeVideoMetadata_EmbeddedMp4_ExtractsResolution()
    {
        var path = CreatePptxWithEmbeddedVideo(BuildMinimalMp4(width: 1920, height: 1080));

        var result = Service.AnalyzeVideoMetadata(path);

        var videoTrack = result.Parts[0].Tracks.First(t => t.TrackType == "video");
        Assert.Equal(1920, videoTrack.Width);
        Assert.Equal(1080, videoTrack.Height);
    }

    [Fact]
    public void AnalyzeVideoMetadata_EmbeddedMp4_ExtractsDuration()
    {
        // 5 seconds: duration=450000, timescale=90000
        var path = CreatePptxWithEmbeddedVideo(BuildMinimalMp4(timescale: 90000, duration: 450000));

        var result = Service.AnalyzeVideoMetadata(path);

        var videoTrack = result.Parts[0].Tracks.First(t => t.TrackType == "video");
        Assert.NotNull(videoTrack.DurationSeconds);
        Assert.InRange(videoTrack.DurationSeconds.Value, 4.9, 5.1);
    }

    [Fact]
    public void AnalyzeVideoMetadata_EmbeddedMp4_ExtractsBitrate()
    {
        var path = CreatePptxWithEmbeddedVideo();

        var result = Service.AnalyzeVideoMetadata(path);

        var videoTrack = result.Parts[0].Tracks.First(t => t.TrackType == "video");
        Assert.NotNull(videoTrack.Bitrate);
        Assert.True(videoTrack.Bitrate > 0);
    }

    [Fact]
    public void AnalyzeVideoMetadata_EmbeddedMp4_PartHasPositiveFileSize()
    {
        var path = CreatePptxWithEmbeddedVideo();

        var result = Service.AnalyzeVideoMetadata(path);

        Assert.All(result.Parts, p => Assert.True(p.FileSizeBytes > 0));
    }

    [Fact]
    public void AnalyzeVideoMetadata_EmbeddedMp4_PartHasContentType()
    {
        var path = CreatePptxWithEmbeddedVideo();

        var result = Service.AnalyzeVideoMetadata(path);

        Assert.All(result.Parts, p =>
            Assert.Contains("video", p.ContentType, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AnalyzeVideoMetadata_EmbeddedMp4_PartHasNoError()
    {
        var path = CreatePptxWithEmbeddedVideo();

        var result = Service.AnalyzeVideoMetadata(path);

        Assert.All(result.Parts, p => Assert.Null(p.Error));
    }

    [Fact]
    public void AnalyzeVideoMetadata_EmbeddedMp4_TotalTracksMatchesSumOfPartTracks()
    {
        var path = CreatePptxWithEmbeddedVideo();

        var result = Service.AnalyzeVideoMetadata(path);

        var expected = result.Parts.Sum(p => p.Tracks.Count);
        Assert.Equal(expected, result.TotalTracks);
    }

    [Fact]
    public void AnalyzeVideoMetadata_CustomResolution_ExtractsCorrectDimensions()
    {
        var mp4 = BuildMinimalMp4(width: 3840, height: 2160);
        var path = CreatePptxWithEmbeddedVideo(mp4);

        var result = Service.AnalyzeVideoMetadata(path);

        var videoTrack = result.Parts[0].Tracks.First(t => t.TrackType == "video");
        Assert.Equal(3840, videoTrack.Width);
        Assert.Equal(2160, videoTrack.Height);
    }

    // ── Tests: Non-MP4 media types ─────────────────────────────────────

    [Fact]
    public void AnalyzeVideoMetadata_NonMp4Media_HandlesGracefully()
    {
        // Embed random bytes as "video/mp4" — should fail to parse but not throw
        var randomBytes = new byte[128];
        new Random(42).NextBytes(randomBytes);
        var path = CreatePptxWithEmbeddedVideo(randomBytes);

        var result = Service.AnalyzeVideoMetadata(path);

        Assert.True(result.Success);
        Assert.Equal(1, result.VideoPartsFound);
        var part = Assert.Single(result.Parts);
        Assert.NotNull(part.Error);
        Assert.Empty(part.Tracks);
    }

    [Fact]
    public void AnalyzeVideoMetadata_OnlyImages_ReturnsNoParts()
    {
        // PPTX with images but no video
        var path = CreatePptxWithSlides(
            new TestSlideDefinition { IncludeImage = true });

        var result = Service.AnalyzeVideoMetadata(path);

        Assert.True(result.Success);
        Assert.Equal(0, result.VideoPartsFound);
    }

    // ── Tests: Audio tracks ────────────────────────────────────────────

    [Fact]
    public void AnalyzeVideoMetadata_EmbeddedAudio_FindsAudioTrack()
    {
        var path = CreatePptxWithEmbeddedAudio();

        var result = Service.AnalyzeVideoMetadata(path);

        Assert.True(result.Success);
        Assert.Equal(1, result.VideoPartsFound);
        var part = Assert.Single(result.Parts);
        Assert.Contains(part.Tracks, t => t.TrackType == "audio");
    }

    [Fact]
    public void AnalyzeVideoMetadata_AudioTrack_ExtractsCodec()
    {
        var path = CreatePptxWithEmbeddedAudio();

        var result = Service.AnalyzeVideoMetadata(path);

        var audioTrack = result.Parts[0].Tracks.First(t => t.TrackType == "audio");
        Assert.Equal("mp4a", audioTrack.Codec);
    }

    [Fact]
    public void AnalyzeVideoMetadata_AudioTrack_HasNoResolution()
    {
        var path = CreatePptxWithEmbeddedAudio();

        var result = Service.AnalyzeVideoMetadata(path);

        var audioTrack = result.Parts[0].Tracks.First(t => t.TrackType == "audio");
        Assert.Null(audioTrack.Width);
        Assert.Null(audioTrack.Height);
    }

    [Fact]
    public void AnalyzeVideoMetadata_AudioTrack_ExtractsChannelCount()
    {
        var path = CreatePptxWithEmbeddedAudio(BuildMinimalAudioMp4(channelCount: 2));

        var result = Service.AnalyzeVideoMetadata(path);

        var audioTrack = result.Parts[0].Tracks.First(t => t.TrackType == "audio");
        Assert.Equal(2, audioTrack.ChannelCount);
    }

    [Fact]
    public void AnalyzeVideoMetadata_AudioTrack_ExtractsSampleRate()
    {
        var path = CreatePptxWithEmbeddedAudio(BuildMinimalAudioMp4(sampleRate: 44100));

        var result = Service.AnalyzeVideoMetadata(path);

        var audioTrack = result.Parts[0].Tracks.First(t => t.TrackType == "audio");
        Assert.Equal(44100, audioTrack.SampleRate);
    }

    // ── Tests: File not found ──────────────────────────────────────────

    [Fact]
    public void AnalyzeVideoMetadata_FileNotFound_ThrowsException()
    {
        var nonExistentPath = Path.Join(Path.GetTempPath(), "nonexistent_" + Guid.NewGuid() + ".pptx");

        Assert.ThrowsAny<Exception>(() => Service.AnalyzeVideoMetadata(nonExistentPath));
    }

    // ── Tests: Message content ─────────────────────────────────────────

    [Fact]
    public void AnalyzeVideoMetadata_WithVideo_MessageContainsPartCount()
    {
        var path = CreatePptxWithEmbeddedVideo();

        var result = Service.AnalyzeVideoMetadata(path);

        Assert.Contains("1 media part", result.Message);
    }
}
