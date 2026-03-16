# Tool Reference: pptx-mcp

All tools operate on `.pptx` files via the `filePath` parameter. Paths can be absolute or relative to the working directory where the MCP server is running.

---

## `pptx_list_slides`

List all slides in a presentation with metadata (index, title, layout).

| Parameter | Type | Required | Description |
|---|---|---|---|
| `filePath` | string | ✓ | Path to the `.pptx` file |

**Example prompt:**
```
List all slides in /path/to/deck.pptx
```

---

## `pptx_list_layouts`

List all slide layouts available in the presentation's template.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `filePath` | string | ✓ | Path to the `.pptx` file |

**Example prompt:**
```
What slide layouts are available in /path/to/deck.pptx?
```

---

## `pptx_get_slide_content`

Extract structured content from a slide: shapes, text, tables, and position/size metadata. Returns JSON. **Prefer this over `pptx_get_slide_xml` when you need to read or reason about slide content.**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `filePath` | string | ✓ | Path to the `.pptx` file |
| `slideIndex` | int | ✓ | Zero-based slide index |

**Example prompt:**
```
Get the content of slide 2 in /path/to/deck.pptx
```

---

## `pptx_get_slide_xml`

Get the raw OpenXML markup for a slide. Useful for debugging or advanced inspection.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `filePath` | string | ✓ | Path to the `.pptx` file |
| `slideIndex` | int | ✓ | Zero-based slide index |

**Example prompt:**
```
Show me the raw XML for slide 0 in /path/to/deck.pptx
```

---

## `pptx_add_slide`

Add a new slide to a presentation using an optional layout name.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `filePath` | string | ✓ | Path to the `.pptx` file |
| `layoutName` | string | — | Name of the layout to use. Defaults to the first available layout. |

**Example prompt:**
```
Add a slide using the "Title and Content" layout to /path/to/deck.pptx
```

---

## `pptx_update_text`

Update the text in a placeholder on a specific slide.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `filePath` | string | ✓ | Path to the `.pptx` file |
| `slideIndex` | int | ✓ | Zero-based slide index |
| `placeholderIndex` | int | ✓ | Zero-based placeholder index on the slide |
| `text` | string | ✓ | New text content |

**Example prompt:**
```
Update the title of slide 1 in /path/to/deck.pptx to "Q2 Results"
```

---

## `pptx_insert_image`

Embed an image onto a slide. Position and size are specified in [EMUs (English Metric Units)](https://docs.microsoft.com/en-us/office/open-xml/measurement-units). 914400 EMUs = 1 inch.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `filePath` | string | ✓ | Path to the `.pptx` file |
| `slideIndex` | int | ✓ | Zero-based slide index |
| `imagePath` | string | ✓ | Path to the image file (`.png`, `.jpg`, `.gif`, `.bmp`) |
| `x` | long | — | Left offset in EMUs. Default: `0` |
| `y` | long | — | Top offset in EMUs. Default: `0` |
| `width` | long | — | Image width in EMUs. Default: `2743200` (~3 inches) |
| `height` | long | — | Image height in EMUs. Default: `2057400` (~2.25 inches) |

**Example prompt:**
```
Insert /path/to/chart.png on slide 3 of /path/to/deck.pptx
```

---

## Notes

- All slide and placeholder indices are **zero-based**.
- Tools that read data (`pptx_list_slides`, `pptx_get_slide_content`, etc.) are safe to call repeatedly—they don't modify the file.
- Tools that write (`pptx_add_slide`, `pptx_update_text`, `pptx_insert_image`) modify the file in place. Make a backup before experimenting.
