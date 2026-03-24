# pptx-mcp

Give AI agents native access to PowerPoint. Read slides, extract content, add slides, update text and shape data, and insert images—all through natural language, without touching Office.

**pptx-mcp** is a .NET [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server that bridges AI reasoning and PowerPoint files. It's built for developers and power users who want to automate content extraction, data-driven slide updates, and intelligent presentation generation.

---

## Quick Install

> **Note:** NuGet publishing is planned. For now, build from source.

```bash
git clone https://github.com/jongalloway/pptx-mcp.git
cd pptx-mcp
dotnet build PptxMcp.slnx --configuration Release
```

### Wire it up to Claude Desktop

Add the following to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "pptx-mcp": {
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "/absolute/path/to/pptx-mcp/src/PptxMcp",
        "--configuration",
        "Release"
      ]
    }
  }
}
```

Once NuGet publishing is set up, this simplifies to:

```json
{
  "mcpServers": {
    "pptx-mcp": {
      "command": "pptx-mcp"
    }
  }
}
```

See [docs/QUICKSTART.md](docs/QUICKSTART.md) for a full walkthrough.

---

## What You Can Do

### Tools

| Tool | What it does |
|---|---|
| `pptx_list_slides` | List all slides with metadata |
| `pptx_list_layouts` | List available slide layouts |
| `pptx_get_slide_content` | Extract structured content from a slide (shapes, text, tables) |
| `pptx_get_slide_xml` | Get the raw XML for a slide (power users) |
| `pptx_add_slide` | Add a new slide using a named layout |
| `pptx_add_slide_from_layout` | Create a slide from a named layout and populate placeholders by semantic key |
| `pptx_duplicate_slide` | Clone a slide, including related parts, with optional placeholder overrides |
| `pptx_update_text` | Update text in a placeholder on a slide by index |
| `pptx_update_slide_data` | Update a named or indexed shape while preserving formatting — preferred for single data-driven updates |
| `pptx_batch_update` | Apply multiple named text updates across a deck in one open/save cycle |
| `pptx_insert_image` | Embed an image (PNG, JPG, GIF) on a slide |
| `pptx_replace_image` | Replace an image in an existing picture shape — inherits geometry from the layout, no manual coordinates needed |
| `pptx_insert_table` | Insert a new table onto a slide with headers and data rows |
| `pptx_update_table` | Update cell values in an existing table — target by name or zero-based index |
| `pptx_chart_data` | Read or update data in an existing chart (bar, column, line, pie, area, scatter) — preserves all chart formatting |
| `pptx_write_notes` | Set or replace speaker notes on a slide (supports append and multi-paragraph) |
| `pptx_move_slide` | Move a slide to a different position |
| `pptx_delete_slide` | Remove a slide by its 1-based slide number |
| `pptx_reorder_slides` | Batch reorder all slides by providing the new sequence |
| `pptx_extract_talking_points` | Extract the highest-signal talking points from each slide |
| `pptx_export_markdown` | Export a full presentation to a structured markdown file |
| `pptx_analyze_file_size` | Analyze file size breakdown by category (slides, images, video/audio, masters, layouts) |
| `pptx_analyze_media` | List and analyze all media assets (images, video, audio) with duplicate detection |
| `pptx_find_unused_layouts` | Find unused slide masters and layouts with estimated space savings |
| `pptx_remove_unused_layouts` | Remove unused slide layouts and orphaned masters with before/after validation |
| `pptx_deduplicate_media` | Deduplicate identical media by hash, redirect references, remove orphaned copies |
| `pptx_optimize_images` | Compress/optimize images by downscaling, format conversion, and recompression |

**When to use `pptx_update_slide_data` vs `pptx_update_text`:** Use `pptx_update_slide_data` when shapes have descriptive names (check `pptx_get_slide_content`) — it targets shapes by name and preserves their existing formatting. Use `pptx_update_text` for anonymous placeholders identified only by index.

**When to use `pptx_add_slide_from_layout`:** Use `pptx_add_slide_from_layout` when you want PowerPoint to respect an existing template layout while you fill placeholders in one call. Placeholder keys use semantic names like `Title`, `Body:1`, or `Picture:2`.

**When to use `pptx_duplicate_slide`:** Use `pptx_duplicate_slide` when you already have a styled slide you want to reuse. It deep-clones the slide and related parts, then applies optional placeholder overrides to the duplicate only.

**When to use `pptx_replace_image`:** Use `pptx_replace_image` to swap the image in an existing picture shape (by name or index). The shape's position and size are preserved from the layout, so no EMU coordinates are needed. Supports PNG, JPEG, and SVG. Use the optional `altText` parameter for accessibility.

**When to use `pptx_batch_update`:** Use `pptx_batch_update` when you already know several shape names and want to refresh an entire deck in one pass. It applies multiple text mutations in one open/save cycle and returns per-mutation success details.

**When to use `pptx_insert_table`:** Use `pptx_insert_table` to add a new DrawingML table to a slide. Pass column headers and row data as arrays. Position and size are specified in EMUs (914,400 EMUs = 1 inch); defaults place a full-width table 1.5 inches from the top. Assign a name via `tableName` so you can target the table later with `pptx_update_table`.

**When to use `pptx_update_table`:** Use `pptx_update_table` to overwrite cell values in an existing table. Locate the table by its `tableName` (case-insensitive, takes precedence) or by `tableIndex` (zero-based). Each update targets a cell by zero-based `row` and `column` indices. Out-of-range updates are silently skipped and counted in `CellsSkipped`.

**When to use `pptx_chart_data`:** Use `pptx_chart_data` to read or refresh data in an existing chart without touching its styling. Call with `action: Read` to inspect chart type, series names, categories, and values. Call with `action: Update` to replace values — provide an `updates` array where each entry specifies the zero-based `SeriesIndex` and the new `Values`, `Categories`, and/or `SeriesName`. All chart formatting (colors, fonts, line styles) is preserved. Supports Column, Bar, Line, Pie, Area, Scatter, and their 3D/Doughnut variants. Locate the chart by `chartName` or `chartIndex`; if the slide has only one chart, both may be omitted.

### Resources

Resources let agents browse presentation state without imperative tool calls. Access them using URIs of the form `pptx://{file}/{resource}` where `{file}` is the URL-encoded absolute path to the `.pptx` file.

| Resource URI | What it returns |
|---|---|
| `pptx://{file}/slides` | JSON array of all slides with index, title, notes, and placeholder count |
| `pptx://{file}/layouts` | JSON array of available slide layouts with index and name |
| `pptx://{file}/shape-map` | JSON object keyed by zero-based slide index (e.g. `"0"`, `"1"`), each containing all named shapes with type, placeholder type, and current text |

**Example:** To browse slide content in `/home/user/deck.pptx`, access `pptx://%2Fhome%2Fuser%2Fdeck.pptx/slides` (file path URL-encoded).

The `shape-map` resource is especially useful for discovering shape names before calling `pptx_update_slide_data`.

### Prompts

Prompts are reusable workflow templates that give agents a structured starting point for common tasks.

| Prompt | What it does |
|---|---|
| `refresh-qbr-deck` | Step-by-step workflow for refreshing a QBR deck with live metrics from a data source |
| `create-agenda-slide` | Adds an agenda slide listing current slide titles at the end of the deck |
| `replace-kpi-placeholders` | Finds all placeholder tokens (e.g. `{{KPI_NAME}}`, `[VALUE]`, `TBD`) and replaces them with real values |

### Completions

pptx-mcp supports argument auto-completion for:
- **`layoutName`** — autocompletes layout names from the actual presentation file (requires `file` or `filePath` context argument)
- **`shapeName`** — autocompletes shape names across all slides in a single file pass (requires `file` or `filePath` context argument)  
- **`placeholderType`** — suggests standard OpenXML placeholder type names (`title`, `body`, `ctrTitle`, etc.)

**Limitations:** pptx-mcp updates text content, inserts images, creates tables, and refreshes data in existing charts. It does not create charts from scratch or modify slide master/theme styles. Complex layout changes should be done in PowerPoint directly.

→ Full parameter docs and examples: [docs/TOOL_REFERENCE.md](docs/TOOL_REFERENCE.md)

---

## Use Cases

**Meeting Prep Assistant**
Ask your AI assistant to read your keynote deck, pull out the key talking points per slide, and give you a quick summary before you walk into the room. Use `pptx_extract_talking_points` to get ranked, noise-filtered bullet points from every slide in a single call.

**Documentation Generator**
Export a training presentation to markdown for your knowledge base. Use `pptx_export_markdown` to convert the whole deck—headings, bullets, tables, and image references—to a structured `.md` file in one step. Keep the deck as the source of truth; let the agent keep the docs in sync.

**Data Dashboard Updater**
Connect pptx-mcp with a data source MCP. Your agent fetches today's KPIs and updates the metrics slide automatically—no manual editing needed. Use the `refresh-qbr-deck` prompt to guide the workflow, or browse the `pptx://{file}/shape-map` resource first to discover shape names. See the [multi-source composition guide](docs/MULTI_SOURCE_COMPOSITION.md) and the included [mock-data-mcp example](examples/mock-data-mcp/) to run it locally.

See [docs/EXAMPLES.md](docs/EXAMPLES.md) for complete agent prompts and sample tool call sequences for each scenario.

---

## Contributing / Dev Setup

Requirements: [.NET 10 SDK](https://dotnet.microsoft.com/download)

```bash
# Build
dotnet build PptxMcp.slnx --configuration Release

# Test
dotnet test PptxMcp.slnx --configuration Release
```

Architecture overview and internal docs live in [docs/PRD.md](docs/PRD.md).

Contributions are welcome—open an issue or a PR.