# Shape Resolution Guide

When updating shapes in PowerPoint presentations via pptx-tools, you need to identify them clearly. This guide explains the three methods for targeting shapes and how to choose the right one for your workflow.

---

## Overview: Three Ways to Target Shapes

| Method | Example | Use Case | Speed |
|--------|---------|----------|-------|
| **By Name** | `shapeName: "KPI_Metric1"` | Named shapes in templates | Fast (direct lookup) |
| **By Index** | `shapeIndex: 2` (slide-scoped, 0-based) | Programmatic iteration or discovery | Very fast |
| **By Placeholder Type** | `placeholderType: "Body"` or `"Title"` | Standard slide layouts | Medium (layout-aware) |

---

## Method 1: Shape Names (Recommended for Templates)

### What's a Shape Name?

Every shape in a PowerPoint presentation has a **unique name** within its slide. Names are human-readable and set by the template author or programmatically via tools. Example names: `"Title"`, `"Content"`, `"Chart 1"`, `"Logo"`.

### How to Discover Shape Names

**Option A: Use the `pptx://{file}/shape-map` Resource**

Before calling any update tool, browse the shape map resource to see all available shapes:

```
Resource URI: pptx://{file}/shape-map
```

This returns a JSON object keyed by slide index:

```json
{
  "0": {
    "Title": {
      "type": "TextShape",
      "placeholderType": "Title",
      "text": "Welcome to pptx-tools"
    },
    "Content": {
      "type": "TextShape",
      "placeholderType": "Body",
      "text": "Learn how to automate PowerPoint"
    }
  },
  "1": {
    "Chart 1": {
      "type": "Chart",
      "placeholderType": null,
      "text": "[chart data not extracted]"
    }
  }
}
```

**Option B: Call `pptx_get_slide_content`**

Extract detailed content from a specific slide:

```json
{
  "name": "pptx_get_slide_content",
  "arguments": {
    "filePath": "/path/to/presentation.pptx",
    "slideNumber": 1
  }
}
```

Returns all shapes, images, tables, and their names.

### Using Shape Names in Tools

Once you know the name, pass it to tools like `pptx_update_slide_data`:

```json
{
  "name": "pptx_update_slide_data",
  "arguments": {
    "filePath": "/presentations/qbr.pptx",
    "slideNumber": 2,
    "shapeName": "KPI_Revenue",
    "newText": "$5.2M (↑12%)"
  }
}
```

### Naming Conventions in pptx-tools

- **Case-sensitive by default** — `"Title"` ≠ `"title"`
- **Layout shapes** — Standard layouts use names like `"Title"`, `"Content"`, `"Subtitle"` (match your slide master)
- **Custom shapes** — When creating slides, assign semantic names: `"KPI_Metric1"`, `"Chart_Sales"`, `"Logo"`
- **Best practice** — Use snake_case or PascalCase for custom shapes to avoid conflicts with layout defaults

---

## Method 2: Shape Index (Zero-Based, Slide-Scoped)

### What's a Shape Index?

Each slide has an ordered list of shapes. Shapes are indexed **0-based starting from the back** (lowest Z-order) to the front (highest Z-order). The index is unique **within the slide only** — different slides can have a shape at index 0.

### How Indexing Works

If a slide contains 5 shapes (title, content, background, image, footer), they are indexed 0–4 in rendering order:

```
Shape Index 0: Background rectangle
Shape Index 1: Title (front to back)
Shape Index 2: Content box
Shape Index 3: Inserted image
Shape Index 4: Footer (topmost)
```

### When to Use Shape Index

- **Batch operations** — iterate through all shapes on a slide
- **Programmatic workflows** — when you don't know shape names in advance
- **Fallback** — when a shape lacks a meaningful name
- **One-off updates** — when you just need to update the 3rd shape without caring about its name

### Using Shape Index in Tools

```json
{
  "name": "pptx_update_slide_data",
  "arguments": {
    "filePath": "/presentations/deck.pptx",
    "slideNumber": 1,
    "shapeIndex": 2,
    "newText": "Updated via index"
  }
}
```

### Discovering Shape Indices

Use `pptx_get_slide_content` and count the shapes in returned order:

```json
{
  "name": "pptx_get_slide_content",
  "arguments": {
    "filePath": "/presentations/deck.pptx",
    "slideNumber": 1
  }
}
```

Response (shapes listed in order):
```json
{
  "slideNumber": 1,
  "title": "Quarterly Review",
  "shapes": [
    { "index": 0, "name": "Background", "type": "Rectangle", "text": "" },
    { "index": 1, "name": "Title", "type": "TextShape", "text": "Quarterly Review" },
    { "index": 2, "name": "Content", "type": "TextShape", "text": "Key metrics..." }
  ]
}
```

---

## Method 3: Placeholder Types (For Layout-Based Slides)

### What's a Placeholder Type?

Placeholders are special shape slots defined by a slide layout. They are **type-identified** rather than name-identified. Examples: `"Title"`, `"Body"`, `"Picture"`, `"SlideNumber"`.

### Standard Placeholder Types

| Type | Typical Use | Count per Slide |
|------|---|---|
| `Title` | Slide title | 0–1 |
| `Body` | Bullet points / body text | 0–4 |
| `Picture` | Image placeholder | 0–2 |
| `CenterTitle` | Centered title (title slide only) | 0–1 |
| `SubTitle` | Subtitle (title slide) | 0–1 |
| `SlideNumber` | Slide number footer | 0–1 |
| `Datetime` | Date/time footer | 0–1 |
| `Footer` | Footer text | 0–1 |

### When to Use Placeholder Types

- **Template-driven workflows** — creating slides from layouts with known placeholder structure
- **Multi-source updates** — updating the same placeholder across multiple slides
- **Safe positioning** — placeholder shapes are positioned by the layout; no EMU calculations needed
- **Format preservation** — placeholder shapes inherit formatting from the layout theme

### Using Placeholder Types in Tools

When creating slides from a layout, target placeholders by semantic identifier:

```json
{
  "name": "pptx_add_slide_from_layout",
  "arguments": {
    "filePath": "/presentations/qbr.pptx",
    "layoutName": "Title and Content",
    "placeholderValues": {
      "Title": "Q1 Results",
      "Body:1": "Revenue: $2.3M\nGrowth: +18%",
      "Body:2": "Key wins: 3 new enterprise clients"
    }
  }
}
```

The tool populates the layout's `Title` placeholder and multiple `Body` placeholders.

### Placeholder Naming Convention

Placeholders are identified by type, optionally indexed:

- `"Title"` — the main title placeholder (usually 1 per slide)
- `"Body"` or `"Body:1"` — first body placeholder
- `"Body:2"` — second body placeholder (if layout has multiple)
- `"Picture:1"` — first picture placeholder
- `"Picture:2"` — second picture placeholder
- Similar pattern for other types: `"SubTitle"`, `"CenterTitle"`, etc.

---

## Comparison: Which Method to Use?

### Scenario 1: Updating a KPI Dashboard with Named Shapes

**Context:** Your presentation has a slide with pre-named shapes: `"KPI_Revenue"`, `"KPI_Growth"`, `"KPI_Forecast"`.

**Best method:** Shape Name
```json
{
  "name": "pptx_update_slide_data",
  "arguments": {
    "filePath": "/presentations/dashboard.pptx",
    "slideNumber": 1,
    "shapeName": "KPI_Revenue",
    "newText": "$5.2M"
  }
}
```
**Why:** Names are semantic and descriptive. No guessing about order or layout quirks.

---

### Scenario 2: Batch-Updating All Shapes on a Slide

**Context:** You want to programmatically update several shapes without knowing their names.

**Best method:** Shape Index
```json
{
  "name": "pptx_get_slide_content",
  "arguments": {
    "filePath": "/presentations/deck.pptx",
    "slideNumber": 1
  }
}
```

Response tells you which shapes exist and their indices. Then loop:
```json
{
  "name": "pptx_update_slide_data",
  "arguments": {
    "filePath": "/presentations/deck.pptx",
    "slideNumber": 1,
    "shapeIndex": 1,
    "newText": "Updated content"
  }
}
```

**Why:** Index is reliable and works on any slide, regardless of layout or naming.

---

### Scenario 3: Creating Slides from a Company Template

**Context:** Your template has a standard `"Title and Content"` layout with known placeholders.

**Best method:** Placeholder Type
```json
{
  "name": "pptx_add_slide_from_layout",
  "arguments": {
    "filePath": "/presentations/company-deck.pptx",
    "layoutName": "Title and Content",
    "placeholderValues": {
      "Title": "New Feature Launch",
      "Body:1": "• Released in Q2\n• 3 major improvements"
    }
  }
}
```

**Why:** Placeholders are layout-aware, format-preserving, and semantic. The layout handles positioning automatically.

---

### Scenario 4: One-Off Shape Update Without Discovery

**Context:** You don't want to call `pptx_get_slide_content` first; you just want to update something quickly.

**Best method:** Shape Index or Name (with fallback)

Try by name first (fast if name is known):
```json
{
  "name": "pptx_update_slide_data",
  "arguments": {
    "filePath": "/presentations/deck.pptx",
    "slideNumber": 2,
    "shapeName": "Title"
  }
}
```

If that fails (shape name unknown), use index:
```json
{
  "name": "pptx_update_slide_data",
  "arguments": {
    "filePath": "/presentations/deck.pptx",
    "slideNumber": 2,
    "shapeIndex": 1
  }
}
```

**Why:** Minimize API calls by trying the most likely method first.

---

## Troubleshooting Shape Resolution

### Shape Not Found by Name

**Problem:** `pptx_update_slide_data` with `shapeName: "KPI_Metric1"` returns error.

**Check:**
1. Is the shape name spelled correctly? (Case-sensitive)
2. Is it on the correct slide? (Check `slideNumber`)
3. Use `pptx_get_slide_content` to confirm the shape exists and its exact name
4. Look for special characters or spaces in the name

**Solution:** Use `pptx://{file}/shape-map` resource to discover the actual shape name, then retry.

---

### Shape Index Out of Range

**Problem:** `pptx_update_slide_data` with `shapeIndex: 5` on a slide with only 3 shapes.

**Check:**
1. How many shapes does the slide actually have? (Call `pptx_get_slide_content`)
2. Is the index 0-based? (0, 1, 2, ... not 1, 2, 3)

**Solution:** Use 0-based indexing and verify the shape count before calling the tool.

---

### Placeholder Not Populated in Created Slide

**Problem:** You called `pptx_add_slide_from_layout` with `"Body:1": "..."` but the text doesn't appear.

**Checks:**
1. Does the layout actually have a Body placeholder? (Call `pptx_list_layouts` then inspect in PowerPoint)
2. Is the placeholder identifier correct? (Try `"Body"` instead of `"Body:1"`, or vice versa)
3. Is the text too long for the placeholder? (Try shorter text)

**Solution:** Use `pptx_get_slide_content` on the created slide to see which placeholders were actually populated. Adjust the `placeholderValues` object accordingly.

---

## Related Resources

- [TOOL_REFERENCE.md](TOOL_REFERENCE.md) — Full tool documentation with all parameter options
- [EMU_GUIDE.md](EMU_GUIDE.md) — Positioning and sizing with EMU coordinates
- [QUICKSTART.md](QUICKSTART.md) — First-time setup and basic usage
