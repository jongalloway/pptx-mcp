# Skill: OpenXML Package Analysis

**Category:** OpenXML / PPTX Internals  
**Domain:** Media extraction, package introspection, deduplication  
**Maturity:** Established (prior art in MarpToPptx)  
**Complexity:** Medium

---

## Scope

Techniques for analyzing and extracting metadata from PPTX packages at the OPC (Open Packaging Convention) and media levels:

- Enumerating parts by content type and URI patterns
- Extracting media (images, video, audio) with relationship resolution
- Computing media hashes for deduplication
- Traversing layout/master relationships
- Zip-level compression analysis

---

## Core Patterns

### 1. Package Part Enumeration

**Entry Point:**
```csharp
using var doc = PresentationDocument.Open(filePath, false); // false = read-only
var presentationPart = doc.PresentationPart;
var package = presentationPart.OpenXmlPackage.Package;

foreach (var part in package.GetParts())
{
    var uri = part.Uri.ToString();        // e.g., /ppt/slides/slide1.xml
    var contentType = part.ContentType;   // e.g., application/vnd.openxmlformats-officedocument.presentationml.slide+xml
    using var stream = part.GetStream();
    var size = stream.Length;
}
```

**Categorization by Content Type:**
| Category | URI Pattern | ContentType |
|----------|-------------|-------------|
| Slides | `/ppt/slides/slide*.xml` | `presentationml.slide+xml` |
| Media | `/ppt/media/*` | `image/png`, `video/mp4`, `audio/mpeg` |
| Themes | `/ppt/theme/*` | `theme+xml` |
| Layouts | `/ppt/slideLayouts/*` | `slideLayout+xml` |
| Masters | `/ppt/slideMasters/*` | `slideMaster+xml` |
| Rels | `/ppt/**/*.rels` | `relationships+xml` |

### 2. Media Extraction & Relationship Resolution

**Image Extraction (from Picture shapes):**
```csharp
foreach (var slidePart in presentationPart.SlideParts)
{
    var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
    if (shapeTree is null) continue;

    foreach (var picture in shapeTree.Elements<Picture>())
    {
        // Get relationship ID from Blip
        var relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
        if (relationshipId is null) continue;

        // Resolve to ImagePart
        if (slidePart.TryGetPartById(relationshipId, out var part) && part is ImagePart imagePart)
        {
            var contentType = imagePart.ContentType; // "image/png", "image/jpeg", etc.
            using var stream = imagePart.GetStream();
            var bytes = stream.Length;
        }
    }
}
```

**Video/Audio Extraction (from Picture with VideoFromFile/AudioFromFile):**
```csharp
foreach (var picture in shapeTree.Elements<Picture>())
{
    var blip = picture.BlipFill?.Blip;
    
    // Check for video
    var videoFromFile = blip?.GetFirstChild<VideoFromFile>();
    if (videoFromFile?.Embed?.Value is string videoRelId)
    {
        if (slidePart.TryGetPartById(videoRelId, out var part) && part is MediaDataPart videoPart)
        {
            var format = videoPart.ContentType; // "video/mp4"
            var size = videoPart.GetStream().Length;
        }
    }
    
    // Similar for AudioFromFile...
}
```

### 3. Media Hashing for Deduplication

**SHA256 Hashing Pattern (from MarpToPptx):**
```csharp
using System.Security.Cryptography;

private static string ComputeMediaHash(ImagePart imagePart)
{
    using var stream = imagePart.GetStream();
    return Convert.ToHexString(SHA256.HashData(stream));
}

// Usage: identify duplicates
var mediaByHash = new Dictionary<string, ImagePart>();
foreach (var imagePart in allImages)
{
    var hash = ComputeMediaHash(imagePart);
    if (!mediaByHash.ContainsKey(hash))
        mediaByHash[hash] = imagePart; // Canonical
    // else: duplicate found
}
```

### 4. Layout/Master Traversal

**Enumerate all layouts in presentation:**
```csharp
var allLayouts = new Dictionary<string, SlideLayoutPart>();

foreach (var masterPart in presentationPart.SlideMasterParts)
{
    foreach (var layoutPart in masterPart.SlideLayoutParts)
    {
        var layoutId = presentationPart.GetIdOfPart(layoutPart);
        var layoutName = layoutPart.SlideLayout.CommonSlideData?.Name?.Value ?? "Unnamed";
        
        allLayouts[layoutId] = layoutPart;
    }
}
```

**Find usage across all slides:**
```csharp
var usedLayouts = new HashSet<string>();

foreach (var slidePart in presentationPart.SlideParts)
{
    if (slidePart.SlideLayoutPart != null)
    {
        var layoutId = presentationPart.GetIdOfPart(slidePart.SlideLayoutPart);
        usedLayouts.Add(layoutId);
    }
}

var unusedLayouts = allLayouts.Keys.Except(usedLayouts);
```

### 5. Zip-Level Compression Analysis

**Direct ZIP access (if needed for compression ratios):**
```csharp
using System.IO.Compression;

using var zip = ZipFile.OpenRead(filePath);
foreach (var entry in zip.Entries)
{
    var path = entry.FullName;                           // e.g., ppt/slides/slide1.xml
    var uncompressed = entry.Length;                    // Logical size
    var compressed = entry.CompressedLength;            // Disk size
    var ratio = (double)entry.CompressedLength / entry.Length;
    
    // Highly compressible parts (rels, XML) often have ratio < 0.2
    // Binary parts (media) have ratio 0.8–1.0 (already compressed)
}
```

---

## Error Handling & Edge Cases

### Try/Catch for Corrupted PPTX
```csharp
try
{
    using var doc = PresentationDocument.Open(filePath, false);
    // ... analysis ...
}
catch (OpenXmlPackageException ex)
{
    // Malformed package structure
    return new { Success = false, Message = $"Invalid PPTX: {ex.Message}" };
}
catch (InvalidOperationException ex)
{
    // Missing required part (e.g., no PresentationPart)
    return new { Success = false, Message = $"Corrupt presentation: {ex.Message}" };
}
```

### Media in Non-Standard Locations
- **Charts:** Media may be in `/ppt/charts/media/*` (separate enumeration)
- **SmartArt:** OleObjects with embedded media (rare; requires custom XML navigation)
- **Embedded in base64:** Some PPTX variants embed media directly in XML (uncommon)

**Workaround:** Start with slidepart media enumeration; extend if needed.

### Relationship Resolution Failures
```csharp
// Safe pattern: always check TryGetPartById
if (slidePart.TryGetPartById(relationshipId, out var part))
{
    if (part is ImagePart imagePart)
    {
        // Process image
    }
}
else
{
    // Orphaned relationship; may indicate corruption or incomplete save
    log.Warn($"Orphaned relationship: {relationshipId}");
}
```

---

## Common Operations

### Categorize PPTX by Content Size
```csharp
var breakdown = new Dictionary<string, long>
{
    ["slides"] = 0,
    ["media"] = 0,
    ["themes"] = 0,
    ["layouts"] = 0,
    ["masters"] = 0,
    ["other"] = 0,
};

foreach (var part in package.GetParts())
{
    var uri = part.Uri.ToString().ToLower();
    var size = part.GetStream().Length;

    if (uri.Contains("/slides/slide")) breakdown["slides"] += size;
    else if (uri.Contains("/media/")) breakdown["media"] += size;
    // ... etc
}

// Total file size (approximate; accounts for ZIP overhead)
var totalLogical = breakdown.Values.Sum();
var totalDisk = new FileInfo(filePath).Length; // Actual file size
```

### Detect Unused Media
```csharp
var usedMediaParts = new HashSet<PackagePart>();

// Collect all referenced parts
foreach (var slidePart in presentationPart.SlideParts)
{
    var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
    foreach (var picture in shapeTree.Elements<Picture>())
    {
        var relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
        if (relationshipId != null && slidePart.TryGetPartById(relationshipId, out var part))
            usedMediaParts.Add(part);
    }
}

// Find all media parts not in this set
foreach (var part in package.GetParts())
{
    if (part.ContentType.StartsWith("image/") || part.ContentType.StartsWith("video/"))
    {
        if (!usedMediaParts.Contains(part))
            Console.WriteLine($"Unused media: {part.Uri}");
    }
}
```

---

## Testing Notes

**PowerPoint Compatibility Gotchas:**
- Relationship cleanups don't always trigger PowerPoint errors immediately; open in PowerPoint after any mutation
- Some PPTX templates are fragile; deleting parts can cause "repair required" prompts
- Media URLs must remain valid after relationship redirection (use `CreateRelationshipToOtherPart()` API)

**Validation Checklist:**
- [ ] File opens in PowerPoint without warnings
- [ ] All media displays correctly (images, videos)
- [ ] Relationships intact (no "broken link" indicators)
- [ ] File size as expected
- [ ] Round-trip save (open → save → reopen) preserves structure

---

## References

**Prior Art:**
- **MarpToPptx:** `OpenXmlPptxRenderer.cs`, `PptxMarkdownExporter.Media.cs`
- **pptx-mcp:** `PresentationService.cs` (lines 96–111, GetLayouts pattern)

**OpenXML SDK Docs:**
- https://github.com/dotnet/Open-XML-SDK/wiki
- https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document

**Related Skills:**
- `openxml-text-updates` — shape/text mutation patterns
- `openxml-table-ops` — table enumeration (similar relationship patterns)
