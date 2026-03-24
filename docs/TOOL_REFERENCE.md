# pptx-mcp Tool Reference

Complete reference for all MCP tools exposed by pptx-mcp, organized alphabetically.

---

## Table of Contents

- [pptx_add_slide](#pptx_add_slide)
- [pptx_add_slide_from_layout](#pptx_add_slide_from_layout)
- [pptx_analyze_file_size](#pptx_analyze_file_size)
- [pptx_analyze_media](#pptx_analyze_media)
- [pptx_batch_update](#pptx_batch_update)
- [pptx_delete_slide](#pptx_delete_slide)
- [pptx_duplicate_slide](#pptx_duplicate_slide)
- [pptx_export_markdown](#pptx_export_markdown)
- [pptx_extract_talking_points](#pptx_extract_talking_points)
- [pptx_find_unused_layouts](#pptx_find_unused_layouts)
- [pptx_remove_unused_layouts](#pptx_remove_unused_layouts)
- [pptx_get_slide_content](#pptx_get_slide_content)
- [pptx_get_slide_xml](#pptx_get_slide_xml)
- [pptx_insert_image](#pptx_insert_image)
- [pptx_insert_table](#pptx_insert_table)
- [pptx_list_layouts](#pptx_list_layouts)
- [pptx_list_slides](#pptx_list_slides)
- [pptx_move_slide](#pptx_move_slide)
- [pptx_reorder_slides](#pptx_reorder_slides)
- [pptx_update_slide_data](#pptx_update_slide_data)
- [pptx_update_table](#pptx_update_table)
- [pptx_update_text](#pptx_update_text)

---

## Notes on Units

Position and size values use **EMUs (English Metric Units)**. 914,400 EMUs = 1 inch. Common reference values:

| Inches | EMUs |
|--------|------|
| 1 in   | 914,400 |
| 2 in   | 1,828,800 |
| 3 in   | 2,743,200 |
| 10 in  | 9,144,000 |

A standard widescreen (16:9) slide is 9,144,000 × 5,143,500 EMUs. A standard 4:3 slide is 9,144,000 × 6,858,000 EMUs.

---

## pptx_add_slide

**Description:** Add a new slide to a PowerPoint presentation.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `layoutName` | string | ❌ Optional | Name of the slide layout to use. Defaults to the first available layout. Use `pptx_list_layouts` to see available layout names. |

### Returns

A plain-text confirmation message with the zero-based index of the newly created slide.

```
Slide added successfully at index 3.
```

On error:
```
Error: File not found: /path/to/presentation.pptx
```

### Example

**Request:**
```json
{
  "name": "pptx_add_slide",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "layoutName": "Title and Content"
  }
}
```

**Response:**
```
Slide added successfully at index 5.
```

**Request (use default layout):**
```json
{
  "name": "pptx_add_slide",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx"
  }
}
```

**Response:**
```
Slide added successfully at index 5.
```

---

## pptx_add_slide_from_layout

**Description:** Create a new slide from a named layout, keep the slide linked to that layout for template inheritance, and optionally populate placeholders in the same call.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `layoutName` | string | ✅ Required | Exact layout name to use. Call `pptx_list_layouts` first if you need to discover available names. |
| `placeholderValues` | object | ❌ Optional | Object keyed by semantic placeholder identifiers like `Title`, `Body:1`, or `Picture:2`. Values are the replacement text to apply. |
| `insertAt` | integer | ❌ Optional | 1-based insertion position. Defaults to appending the new slide at the end of the presentation. |

### Returns

A JSON result describing the created slide.

```json
{
  "Success": true,
  "SlideNumber": 4,
  "LayoutName": "Title and Content",
  "PlaceholdersPopulated": 2,
  "Message": "Added slide 4 from layout 'Title and Content'."
}
```

### Example

```json
{
  "name": "pptx_add_slide_from_layout",
  "arguments": {
    "filePath": "/presentations/qbr.pptx",
    "layoutName": "Title and Content",
    "placeholderValues": {
      "Title": "Agenda",
      "Body:1": "Wins\nRisks\nNext steps"
    }
  }
}
```

---

## pptx_analyze_file_size

**Description:** Analyze the file size breakdown of a PowerPoint presentation by category. Scans all parts in the PPTX package and reports sizes broken down into slides, images, video/audio, slide masters, slide layouts, and other parts. Each category includes a subtotal and per-part detail.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |

### Returns

Structured JSON with file size breakdown:

```json
{
  "Success": true,
  "FilePath": "/presentations/quarterly-review.pptx",
  "TotalFileSize": 2458624,
  "TotalPartSize": 3145728,
  "Categories": [
    {
      "Name": "slides",
      "TotalSize": 45320,
      "PartCount": 3,
      "Parts": [
        { "Path": "/ppt/slides/slide1.xml", "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml", "Size": 15107 }
      ]
    },
    {
      "Name": "images",
      "TotalSize": 2890000,
      "PartCount": 5,
      "Parts": []
    }
  ],
  "Message": "Analyzed 42 parts across 6 categories."
}
```

On error, returns the same structure with `Success: false` and all categories present with zero totals.

### Example

**Request:**
```json
{
  "name": "pptx_analyze_file_size",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx"
  }
}
```

---

## pptx_analyze_media

**Description:** List and analyze all media assets (images, video, audio) in a PowerPoint presentation. For each media part: reports name, content type, size, SHA256 hash, and which slides reference it. Detects duplicate media (same content hash) and groups them for deduplication analysis.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |

### Returns

Structured JSON with media analysis:

```json
{
  "Success": true,
  "FilePath": "/presentations/quarterly-review.pptx",
  "TotalMediaCount": 5,
  "TotalMediaSize": 2890000,
  "DuplicateGroupCount": 1,
  "DuplicateSavingsBytes": 450000,
  "MediaParts": [
    {
      "Path": "/ppt/media/image1.png",
      "ContentType": "image/png",
      "SizeBytes": 450000,
      "Hash": "A1B2C3...",
      "ReferencedBySlides": [1, 3]
    }
  ],
  "DuplicateGroups": [
    {
      "Hash": "A1B2C3...",
      "ContentType": "image/png",
      "SizeBytes": 450000,
      "Parts": ["/ppt/media/image1.png", "/ppt/media/image3.png"],
      "ReferencedBySlides": [1, 3]
    }
  ],
  "Message": "Found 5 media assets (1 duplicate group, 450000 bytes recoverable)."
}
```

On error, returns the same structure with `Success: false` and empty lists.

### Example

**Request:**
```json
{
  "name": "pptx_analyze_media",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx"
  }
}
```

---

## pptx_batch_update

**Description:** Apply many named text updates across multiple slides in a single presentation open/save cycle. Each mutation targets a 1-based `slideNumber` and exact `shapeName`, preserves the shape's existing formatting, and reports per-mutation success or failure.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `mutations` | array<object> | ✅ Required | Array of text mutations. Each item must include `slideNumber` (1-based), `shapeName`, and `newValue`. |

### Returns

A JSON result summarizing how many mutations succeeded or failed, with per-mutation details in request order.

```json
{
  "TotalMutations": 3,
  "SuccessCount": 2,
  "FailureCount": 1,
  "Results": [
    {
      "SlideNumber": 1,
      "ShapeName": "Executive Subtitle",
      "Success": true,
      "Error": null,
      "MatchedBy": "shapeName"
    },
    {
      "SlideNumber": 2,
      "ShapeName": "Missing Shape",
      "Success": false,
      "Error": "No text-capable shape named 'Missing Shape' was found. Available shapes: 1:Revenue Value, 2:Gross Margin",
      "MatchedBy": null
    }
  ]
}
```

Successful mutations are kept even if other mutations fail; the tool does not roll back prior updates.

### Example

**Request:**
```json
{
  "name": "pptx_batch_update",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "mutations": [
      {
        "slideNumber": 2,
        "shapeName": "Revenue Value",
        "newValue": "$4.6M"
      },
      {
        "slideNumber": 3,
        "shapeName": "Risk Body",
        "newValue": "Mitigate churn\nFinish automation"
      }
    ]
  }
}
```

---

## pptx_delete_slide

**Description:** Delete a slide from a presentation by its 1-based slide number. The presentation must contain at least two slides; deleting the last remaining slide is not allowed.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `slideNumber` | integer | ✅ Required | 1-based number of the slide to delete. |

### Returns

A plain-text confirmation message on success, or an error message prefixed with `Error:`.

### Examples

Delete the second slide:

```json
{
  "name": "pptx_delete_slide",
  "arguments": {
    "filePath": "/presentations/quarterly.pptx",
    "slideNumber": 2
  }
}
```

---

## pptx_duplicate_slide

**Description:** Duplicate an existing slide, deep-clone related parts such as images, and optionally override placeholders on the duplicate using semantic placeholder keys.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `slideNumber` | integer | ✅ Required | 1-based slide number to duplicate. |
| `placeholderOverrides` | object | ❌ Optional | Object keyed by semantic placeholder identifiers like `Title` or `Body:2`. Values are applied only to the duplicated slide. |
| `insertAt` | integer | ❌ Optional | 1-based insertion position. Defaults to inserting immediately after the source slide. |

### Returns

A JSON result describing the duplicated slide.

```json
{
  "Success": true,
  "NewSlideNumber": 3,
  "ShapesCopied": 6,
  "OverridesApplied": 1,
  "Message": "Duplicated slide 2 to slide 3."
}
```

### Example

```json
{
  "name": "pptx_duplicate_slide",
  "arguments": {
    "filePath": "/presentations/qbr.pptx",
    "slideNumber": 2,
    "placeholderOverrides": {
      "Title": "EMEA Deep Dive",
      "Body:2": "Mitigate churn\nRebalance pipeline"
    }
  }
}
```

---

## pptx_export_markdown

**Description:** Export a PowerPoint presentation to a structured markdown file. Slide titles become headings, bullet paragraphs become list items, tables become markdown tables, and embedded images are saved to a sibling `{name}_images/` directory with relative references in the markdown.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `outputPath` | string | ❌ Optional | Output path for the `.md` file. Defaults to the presentation path with a `.md` extension. |

### Returns

The generated markdown content as a string. The markdown is also written to the output file.

On error:
```
Error: File not found: /path/to/presentation.pptx
```

### Example

**Request:**
```json
{
  "name": "pptx_export_markdown",
  "arguments": {
    "filePath": "/presentations/onboarding-engineering.pptx",
    "outputPath": "/docs/onboarding-engineering.md"
  }
}
```

**Response (markdown string):**
```markdown
# Engineering Onboarding

---
<!-- Slide 0 -->

## Welcome to the Team

Welcome to the engineering team. This guide walks you through your first week
setup and key processes.

---
<!-- Slide 1 -->

## Development Environment Setup

- Install .NET 10 SDK
- Clone the repository: `git clone https://github.com/org/repo`
- Run `dotnet build` to verify setup
- Run `dotnet test` to confirm all tests pass

---
<!-- Slide 2 -->

## Code Review Process

| Step | Owner | SLA |
|------|-------|-----|
| Open PR | Author | — |
| Review assigned | Tech lead | 1 business day |
| Approval + merge | Reviewer | 2 business days |
```

**Request (default output path — saves `onboarding-engineering.md` next to the .pptx):**
```json
{
  "name": "pptx_export_markdown",
  "arguments": {
    "filePath": "/presentations/onboarding-engineering.pptx"
  }
}
```

---

## pptx_extract_talking_points

**Description:** Extract the highest-signal talking points from each slide in a presentation. The tool prioritizes body text and bullet-like content, filters common noise (formatting-only text, presenter notes labels), and returns up to the requested number of points per slide.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `topN` | integer | ❌ Optional | Maximum number of talking points to return per slide. Defaults to `5`. Must be greater than zero. |

### Returns

A JSON array of slide objects. Each object contains:

| Field | Type | Description |
|-------|------|-------------|
| `SlideIndex` | integer | Zero-based slide index. |
| `Title` | string \| null | Slide title, if present. |
| `Points` | string[] | Ranked talking point strings (up to `topN`). |

On error:
```
Error: File not found: /path/to/presentation.pptx
```

### Example

**Request:**
```json
{
  "name": "pptx_extract_talking_points",
  "arguments": {
    "filePath": "/presentations/q2-product-review.pptx",
    "topN": 3
  }
}
```

**Response:**
```json
[
  {
    "SlideIndex": 0,
    "Title": "Q2 Product Review",
    "Points": []
  },
  {
    "SlideIndex": 1,
    "Title": "Revenue Highlights",
    "Points": [
      "Q2 ARR up 18% YoY",
      "EMEA region grew 34%",
      "Net Revenue Retention: 112%"
    ]
  },
  {
    "SlideIndex": 2,
    "Title": "Roadmap Preview",
    "Points": [
      "GA release: Q3 2025",
      "New integrations: Slack, Teams, Notion",
      "Mobile app entering beta"
    ]
  }
]
```

**Request (use default topN = 5):**
```json
{
  "name": "pptx_extract_talking_points",
  "arguments": {
    "filePath": "/presentations/q2-product-review.pptx"
  }
}
```

---

## pptx_find_unused_layouts

**Description:** Find unused slide masters and layouts in a PowerPoint presentation. Enumerates all masters and layouts, cross-references against actual slide usage, and identifies which could be safely removed with estimated space savings.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |

### Returns

Structured JSON with unused layout analysis:

```json
{
  "Success": true,
  "FilePath": "/presentations/quarterly-review.pptx",
  "TotalMasters": 1,
  "TotalLayouts": 11,
  "UnusedMasterCount": 0,
  "UnusedLayoutCount": 8,
  "EstimatedSavingsBytes": 45000,
  "Masters": [
    {
      "Name": "Office Theme",
      "Uri": "/ppt/slideMasters/slideMaster1.xml",
      "SizeBytes": 12500,
      "IsUsed": true,
      "LayoutCount": 11,
      "UsedLayoutCount": 3
    }
  ],
  "Layouts": [
    {
      "Name": "Title Slide",
      "Uri": "/ppt/slideLayouts/slideLayout1.xml",
      "SizeBytes": 4200,
      "IsUsed": true,
      "MasterName": "Office Theme",
      "ReferencedBySlides": [1]
    }
  ],
  "Warnings": [],
  "Message": "Found 8 unused layouts across 1 master."
}
```

On error, returns the same structure with `Success: false` and empty lists.

### Example

**Request:**
```json
{
  "name": "pptx_find_unused_layouts",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx"
  }
}
```

---

## pptx_remove_unused_layouts

**Description:** Remove unused slide layouts and orphaned slide masters from a PowerPoint presentation. When `layoutUris` is omitted, auto-detects and removes all unused layouts. When specific URIs are provided, removes only those (if they are unused). Validates the package with OpenXmlValidator before and after removal.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file to modify. |
| `layoutUris` | string[] | Optional | Array of layout URIs to remove. Omit to auto-detect all unused layouts. |

### Returns

Structured JSON with removal results and validation status:

```json
{
  "Success": true,
  "FilePath": "/presentations/quarterly-review.pptx",
  "RemovedItems": [
    {
      "Name": "Two Content",
      "Uri": "/ppt/slideLayouts/slideLayout4.xml",
      "Type": "layout",
      "SizeBytes": 3800
    },
    {
      "Name": "Unused Master",
      "Uri": "/ppt/slideMasters/slideMaster2.xml",
      "Type": "master",
      "SizeBytes": 12500
    }
  ],
  "LayoutsRemoved": 1,
  "MastersRemoved": 1,
  "BytesSaved": 16300,
  "Validation": {
    "ErrorsBefore": 0,
    "ErrorsAfter": 0,
    "IsValid": true
  },
  "Message": "Removed 1 layout(s) and 1 master(s). Saved approximately 16,300 bytes."
}
```

On error, returns the same structure with `Success: false` and empty lists.

### Example

**Auto-detect and remove all unused layouts:**
```json
{
  "name": "pptx_remove_unused_layouts",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx"
  }
}
```

**Targeted removal of specific layouts:**
```json
{
  "name": "pptx_remove_unused_layouts",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "layoutUris": [
      "/ppt/slideLayouts/slideLayout4.xml",
      "/ppt/slideLayouts/slideLayout7.xml"
    ]
  }
}
```

---

## pptx_get_slide_content

**Description:** Get structured content from a slide: all shapes with their type, position, size, and text.Returns a JSON object with slide dimensions and a shapes array. Prefer this over `pptx_get_slide_xml` when you need to read or reason about slide content.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `slideIndex` | integer | ✅ Required | Zero-based index of the slide. |

### Returns

A JSON object with the following structure:

```json
{
  "SlideIndex": 0,
  "SlideWidthEmu": 9144000,
  "SlideHeightEmu": 5143500,
  "Shapes": [
    {
      "ShapeId": 2,
      "Name": "Title 1",
      "ShapeType": "Text",
      "X": 457200,
      "Y": 274638,
      "Width": 8229600,
      "Height": 1143000,
      "IsPlaceholder": true,
      "PlaceholderType": "Title",
      "PlaceholderIndex": null,
      "Text": "Q1 Results",
      "Paragraphs": ["Q1 Results"],
      "TableRows": null
    }
  ]
}
```

**`ShapeType` values:** `Text`, `Picture`, `Table`, `GraphicFrame`, `Group`, `Connector`

On error:
```
Error: File not found: /path/to/presentation.pptx
```

### Example

**Request:**
```json
{
  "name": "pptx_get_slide_content",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "slideIndex": 0
  }
}
```

**Response:**
```json
{
  "SlideIndex": 0,
  "SlideWidthEmu": 9144000,
  "SlideHeightEmu": 5143500,
  "Shapes": [
    {
      "ShapeId": 2,
      "Name": "Title 1",
      "ShapeType": "Text",
      "X": 457200,
      "Y": 274638,
      "Width": 8229600,
      "Height": 1143000,
      "IsPlaceholder": true,
      "PlaceholderType": "Title",
      "PlaceholderIndex": null,
      "Text": "Q1 Results",
      "Paragraphs": [
        "Q1 Results"
      ],
      "TableRows": null
    },
    {
      "ShapeId": 3,
      "Name": "Content Placeholder 2",
      "ShapeType": "Text",
      "X": 457200,
      "Y": 1600200,
      "Width": 8229600,
      "Height": 3399600,
      "IsPlaceholder": true,
      "PlaceholderType": "Body",
      "PlaceholderIndex": 1,
      "Text": "Revenue up 12%\nNew customers: 340\nChurn rate: 2.1%",
      "Paragraphs": [
        "Revenue up 12%",
        "New customers: 340",
        "Churn rate: 2.1%"
      ],
      "TableRows": null
    }
  ]
}
```

---

## pptx_get_slide_xml

**Description:** Get the raw XML of a specific slide. Useful for debugging or advanced inspection of slide structure.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `slideIndex` | integer | ✅ Required | Zero-based index of the slide. |

### Returns

A string containing the raw OpenXML markup for the slide part. This is the unprocessed XML as stored inside the .pptx package.

On error:
```
Error: File not found: /path/to/presentation.pptx
```

### Example

**Request:**
```json
{
  "name": "pptx_get_slide_xml",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "slideIndex": 0
  }
}
```

**Response (abbreviated):**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ...>
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          ...
        </p:nvSpPr>
        <p:txBody>
          <a:p><a:r><a:t>Q1 Results</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
```

---

## pptx_insert_image

**Description:** Insert an image onto a slide at a specified position and size.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `slideIndex` | integer | ✅ Required | Zero-based index of the slide. |
| `imagePath` | string | ✅ Required | Absolute or relative path to the image file. Supported formats: `.png`, `.jpg`, `.gif`, `.bmp`. |
| `x` | long | ❌ Optional | Horizontal offset from the left edge of the slide in EMUs. Default: `0`. |
| `y` | long | ❌ Optional | Vertical offset from the top edge of the slide in EMUs. Default: `0`. |
| `width` | long | ❌ Optional | Width of the image in EMUs. Default: `2743200` (~3 inches). |
| `height` | long | ❌ Optional | Height of the image in EMUs. Default: `2057400` (~2.25 inches). |

### Returns

A plain-text confirmation message on success.

```
Image inserted successfully on slide 2.
```

On error:
```
Error: Image file not found: /path/to/image.png
```

### Example

**Request (centered on a 16:9 slide at ~3×2.25 inches):**
```json
{
  "name": "pptx_insert_image",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "slideIndex": 2,
    "imagePath": "/assets/chart.png",
    "x": 3200400,
    "y": 1542700,
    "width": 2743200,
    "height": 2057400
  }
}
```

**Response:**
```
Image inserted successfully on slide 2.
```

**Request (use all defaults — top-left corner at default size):**
```json
{
  "name": "pptx_insert_image",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "slideIndex": 0,
    "imagePath": "/assets/logo.png"
  }
}
```

**Response:**
```
Image inserted successfully on slide 0.
```

---

## pptx_insert_table

**Description:** Insert a new table onto a slide. Pass column headers and data rows as arrays. Creates a DrawingML table (GraphicFrame) with a header row followed by data rows. Position and size are specified in EMUs (English Metric Units); 914,400 EMUs = 1 inch. Assign a `tableName` to make the table targetable by `pptx_update_table`.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `slideNumber` | integer | ✅ Required | 1-based slide number to insert the table on. |
| `headers` | string[] | ✅ Required | Array of column header strings. Determines the number of columns. |
| `rows` | string[][] | ✅ Required | Array of row arrays, each containing cell values for one data row. |
| `tableName` | string | ❌ Optional | Name for the table's GraphicFrame. Defaults to `"Table {id}"`. |
| `x` | long | ❌ Optional | Horizontal offset from the left edge in EMUs. Default: `914400` (1 inch). |
| `y` | long | ❌ Optional | Vertical offset from the top edge in EMUs. Default: `1371600` (1.5 inches). |
| `width` | long | ❌ Optional | Width of the table in EMUs. Default: `7315200` (~8 inches). |
| `height` | long | ❌ Optional | Height of the table in EMUs. Default: `1371600` (1.5 inches). |

### Returns

A JSON object with the following fields:

| Field | Type | Description |
|-------|------|-------------|
| `Success` | boolean | `true` when the table was inserted successfully. |
| `SlideNumber` | integer | 1-based slide number where the table was inserted. |
| `TableName` | string | Name assigned to the table's GraphicFrame. |
| `TableShapeId` | integer | OpenXML shape ID assigned to the table. |
| `TableIndex` | integer | Zero-based index of this table among all tables on the slide. |
| `RowCount` | integer | Number of rows in the inserted table (including the header row). |
| `ColumnCount` | integer | Number of columns in the inserted table. |
| `Message` | string | Human-readable status message. |

### Example

**Request:**
```json
{
  "name": "pptx_insert_table",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "slideNumber": 3,
    "headers": ["Region", "Q1", "Q2", "Q3", "Q4"],
    "rows": [
      ["North", "$1.2M", "$1.4M", "$1.3M", "$1.6M"],
      ["South", "$0.9M", "$1.1M", "$1.0M", "$1.2M"],
      ["West",  "$1.5M", "$1.7M", "$1.8M", "$2.0M"]
    ],
    "tableName": "RegionalSales"
  }
}
```

**Response:**
```json
{
  "Success": true,
  "SlideNumber": 3,
  "TableName": "RegionalSales",
  "TableShapeId": 7,
  "TableIndex": 0,
  "RowCount": 4,
  "ColumnCount": 5,
  "Message": "Table inserted on slide 3 with 4 rows and 5 columns."
}
```

**Response (file not found):**
```json
{
  "Success": false,
  "SlideNumber": 3,
  "TableName": null,
  "TableShapeId": null,
  "TableIndex": null,
  "RowCount": 0,
  "ColumnCount": 0,
  "Message": "File not found: /presentations/quarterly-review.pptx"
}
```

---

## pptx_list_layouts

**Description:** List all available slide layouts in a PowerPoint presentation. Use the returned layout names with `pptx_add_slide`.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |

### Returns

A JSON array of layout objects, each with an `Index` (zero-based) and `Name`.

```json
[
  { "Index": 0, "Name": "Title Slide" },
  { "Index": 1, "Name": "Title and Content" }
]
```

On error:
```
Error: File not found: /path/to/presentation.pptx
```

### Example

**Request:**
```json
{
  "name": "pptx_list_layouts",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx"
  }
}
```

**Response:**
```json
[
  { "Index": 0, "Name": "Title Slide" },
  { "Index": 1, "Name": "Title and Content" },
  { "Index": 2, "Name": "Title Only" },
  { "Index": 3, "Name": "Blank" },
  { "Index": 4, "Name": "Content with Caption" },
  { "Index": 5, "Name": "Picture with Caption" }
]
```

---

## pptx_list_slides

**Description:** List all slides in a PowerPoint presentation with metadata including title, notes, and placeholder count.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |

### Returns

A JSON array of slide objects, each with:

| Field | Type | Description |
|-------|------|-------------|
| `Index` | integer | Zero-based slide index. |
| `Title` | string \| null | Slide title text, if a title placeholder is present. |
| `Notes` | string \| null | Speaker notes text, if any. |
| `PlaceholderCount` | integer | Number of placeholders on the slide. |

On error:
```
Error: File not found: /path/to/presentation.pptx
```

### Example

**Request:**
```json
{
  "name": "pptx_list_slides",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx"
  }
}
```

**Response:**
```json
[
  {
    "Index": 0,
    "Title": "Q1 2025 Business Review",
    "Notes": "Welcome attendees and introduce the agenda.",
    "PlaceholderCount": 2
  },
  {
    "Index": 1,
    "Title": "Revenue Summary",
    "Notes": null,
    "PlaceholderCount": 3
  },
  {
    "Index": 2,
    "Title": null,
    "Notes": "Use this slide to show the product roadmap image.",
    "PlaceholderCount": 1
  }
]
```

---

## pptx_move_slide

**Description:** Move a slide to a different position in the presentation. All other slides shift to fill the gap.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `slideNumber` | integer | ✅ Required | 1-based number of the slide to move. |
| `targetPosition` | integer | ✅ Required | 1-based position to move the slide to. |

### Returns

A plain-text confirmation message on success, or an error message prefixed with `Error:`.

### Examples

Move slide 1 to the end of a 4-slide deck:

```json
{
  "name": "pptx_move_slide",
  "arguments": {
    "filePath": "/presentations/quarterly.pptx",
    "slideNumber": 1,
    "targetPosition": 4
  }
}
```

Move the last slide to the front:

```json
{
  "name": "pptx_move_slide",
  "arguments": {
    "filePath": "/presentations/quarterly.pptx",
    "slideNumber": 4,
    "targetPosition": 1
  }
}
```

---

## pptx_reorder_slides

**Description:** Reorder all slides in a presentation by providing the desired sequence as a 1-based array. Every slide must appear exactly once. Use `pptx_list_slides` first to identify the current slide numbers.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `newOrder` | integer[] | ✅ Required | Array of 1-based slide numbers in the desired order. Must be a permutation of `1..n` where `n` is the total slide count. |

### Returns

A plain-text confirmation message on success, or an error message prefixed with `Error:`.

### Examples

Reverse a 3-slide deck:

```json
{
  "name": "pptx_reorder_slides",
  "arguments": {
    "filePath": "/presentations/quarterly.pptx",
    "newOrder": [3, 2, 1]
  }
}
```

Move the appendix slides (4 and 5) before the content slides (2 and 3), keeping the title slide (1) first:

```json
{
  "name": "pptx_reorder_slides",
  "arguments": {
    "filePath": "/presentations/quarterly.pptx",
    "newOrder": [1, 4, 5, 2, 3]
  }
}
```

---

## pptx_update_slide_data

**Description:** Update text in a named slide shape while preserving the shape's existing formatting. Prefer `shapeName` from `pptx_get_slide_content`; `placeholderIndex` is a zero-based fallback across text-capable slide shapes in slide order.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `slideNumber` | integer | ✅ Required | 1-based slide number to update. |
| `shapeName` | string | ❌ Optional | Shape name to match exactly, ignoring case. When provided and found, it takes precedence over `placeholderIndex`. |
| `placeholderIndex` | integer | ❌ Optional | Zero-based fallback index across text-capable slide shapes on the slide. |
| `newText` | string | ✅ Required | Replacement text for the target shape. Newlines create separate paragraphs. Empty text is allowed. |

### Returns

A JSON result describing whether the update succeeded and which shape was updated.

```json
{
  "Success": true,
  "SlideNumber": 3,
  "RequestedShapeName": "ARR Value",
  "RequestedPlaceholderIndex": null,
  "MatchedBy": "shapeName",
  "ResolvedShapeName": "ARR Value",
  "ResolvedShapeIndex": 2,
  "ResolvedShapeId": 5,
  "PlaceholderType": "body",
  "LayoutPlaceholderIndex": null,
  "PreviousText": "$4.2M",
  "NewText": "$4.6M",
  "Message": "Updated shape 'ARR Value' on slide 3."
}
```

On failure, `Success` is `false` and `Message` explains what was missing or out of range.

### Example

**Request:**
```json
{
  "name": "pptx_update_slide_data",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "slideNumber": 3,
    "shapeName": "ARR Value",
    "newText": "$4.6M"
  }
}
```

**Response:**
```json
{
  "Success": true,
  "SlideNumber": 3,
  "RequestedShapeName": "ARR Value",
  "RequestedPlaceholderIndex": null,
  "MatchedBy": "shapeName",
  "ResolvedShapeName": "ARR Value",
  "ResolvedShapeIndex": 2,
  "ResolvedShapeId": 5,
  "PlaceholderType": "body",
  "LayoutPlaceholderIndex": null,
  "PreviousText": "$4.2M",
  "NewText": "$4.6M",
  "Message": "Updated shape 'ARR Value' on slide 3."
}
```

**Request (fallback by index when no shape name is known):**
```json
{
  "name": "pptx_update_slide_data",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "slideNumber": 3,
    "placeholderIndex": 2,
    "newText": "$4.6M"
  }
}
```

**Response:**
```json
{
  "Success": true,
  "SlideNumber": 3,
  "RequestedShapeName": null,
  "RequestedPlaceholderIndex": 2,
  "MatchedBy": "placeholderIndex",
  "ResolvedShapeName": "ARR Value",
  "ResolvedShapeIndex": 2,
  "ResolvedShapeId": 5,
  "PlaceholderType": "body",
  "LayoutPlaceholderIndex": null,
  "PreviousText": "$4.2M",
  "NewText": "$4.6M",
  "Message": "Updated shape at index 2 on slide 3."
}
```

**Response (shape not found):**
```json
{
  "Success": false,
  "SlideNumber": 3,
  "RequestedShapeName": "Revenue Total",
  "RequestedPlaceholderIndex": null,
  "MatchedBy": null,
  "ResolvedShapeName": null,
  "ResolvedShapeIndex": null,
  "ResolvedShapeId": null,
  "PlaceholderType": null,
  "LayoutPlaceholderIndex": null,
  "PreviousText": null,
  "NewText": "$4.6M",
  "Message": "No shape named 'Revenue Total' found on slide 3, and no placeholderIndex was supplied."
}
```

---

## pptx_update_table

**Description:** Update cell values in an existing table on a slide. Locate the table by name (case-insensitive) or by zero-based index among tables on the slide. Each update targets a specific cell by zero-based row and column indices. Out-of-range cell updates are silently skipped and counted in `CellsSkipped`.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `slideNumber` | integer | ✅ Required | 1-based slide number containing the table. |
| `updates` | object[] | ✅ Required | Array of cell updates. Each object must include `row` (0-based), `column` (0-based), and `value`. |
| `tableName` | string | ❌ Optional | Table name to match (case-insensitive). Takes precedence over `tableIndex`. |
| `tableIndex` | integer | ❌ Optional | Zero-based index among tables on the slide. Used when `tableName` is not provided. |

#### `updates` item shape

| Field | Type | Description |
|-------|------|-------------|
| `row` | integer | Zero-based row index of the cell to update. |
| `column` | integer | Zero-based column index of the cell to update. |
| `value` | string | New text value for the cell. |

### Returns

A JSON object with the following fields:

| Field | Type | Description |
|-------|------|-------------|
| `Success` | boolean | `true` when the update was applied (even if some cells were skipped). |
| `SlideNumber` | integer | 1-based slide number of the updated table. |
| `TableName` | string | Name of the matched table's GraphicFrame. |
| `MatchedBy` | string | How the table was located: `"tableName"` or `"tableIndex"`. |
| `CellsUpdated` | integer | Number of cells that were successfully updated. |
| `CellsSkipped` | integer | Number of update requests skipped because row/column was out of range. |
| `Message` | string | Human-readable status message. |

### Example

**Request:**
```json
{
  "name": "pptx_update_table",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "slideNumber": 3,
    "tableName": "RegionalSales",
    "updates": [
      { "row": 1, "column": 1, "value": "$1.3M" },
      { "row": 1, "column": 2, "value": "$1.5M" },
      { "row": 3, "column": 4, "value": "$2.1M" }
    ]
  }
}
```

**Response:**
```json
{
  "Success": true,
  "SlideNumber": 3,
  "TableName": "RegionalSales",
  "MatchedBy": "tableName",
  "CellsUpdated": 3,
  "CellsSkipped": 0,
  "Message": "Updated 3 cell(s) in table 'RegionalSales' on slide 3."
}
```

**Response (table not found):**
```json
{
  "Success": false,
  "SlideNumber": 3,
  "TableName": "MissingTable",
  "MatchedBy": null,
  "CellsUpdated": 0,
  "CellsSkipped": 0,
  "Message": "No table named 'MissingTable' found on slide 3."
}
```

---

## pptx_update_text

**Description:** Update the text of a placeholder on a slide. Use `pptx_get_slide_content` first to identify the target placeholder index.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `slideIndex` | integer | ✅ Required | Zero-based index of the slide to update. |
| `placeholderIndex` | integer | ✅ Required | Zero-based index of the placeholder on the slide. |
| `text` | string | ✅ Required | New text content for the placeholder. |

### Returns

A plain-text confirmation message on success.

```
Placeholder 1 on slide 0 updated successfully.
```

On error:
```
Error: File not found: /path/to/presentation.pptx
```

### Example

**Request:**
```json
{
  "name": "pptx_update_text",
  "arguments": {
    "filePath": "/presentations/quarterly-review.pptx",
    "slideIndex": 0,
    "placeholderIndex": 0,
    "text": "Q2 2025 Business Review"
  }
}
```

**Response:**
```
Placeholder 0 on slide 0 updated successfully.
```
