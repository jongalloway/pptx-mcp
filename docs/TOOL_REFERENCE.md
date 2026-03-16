# pptx-mcp Tool Reference

Complete reference for all MCP tools exposed by pptx-mcp, organized alphabetically.

---

## Table of Contents

- [pptx_add_slide](#pptx_add_slide)
- [pptx_get_slide_content](#pptx_get_slide_content)
- [pptx_get_slide_xml](#pptx_get_slide_xml)
- [pptx_insert_image](#pptx_insert_image)
- [pptx_list_layouts](#pptx_list_layouts)
- [pptx_list_slides](#pptx_list_slides)
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

## pptx_get_slide_content

**Description:** Get structured content from a slide: all shapes with their type, position, size, and text. Returns a JSON object with slide dimensions and a shapes array. Prefer this over `pptx_get_slide_xml` when you need to read or reason about slide content.

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

---

## Planned Tools

The following tools are planned for Phase 1 and will be added to this reference once implemented:

- **`pptx_extract_talking_points`** — Extract key talking points from one or more slides.
- **`pptx_export_markdown`** — Export presentation content to structured markdown.
