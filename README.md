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

| Tool | What it does |
|---|---|
| `pptx_list_slides` | List all slides with metadata |
| `pptx_list_layouts` | List available slide layouts |
| `pptx_get_slide_content` | Extract structured content from a slide (shapes, text, tables) |
| `pptx_get_slide_xml` | Get the raw XML for a slide (power users) |
| `pptx_add_slide` | Add a new slide using a named layout |
| `pptx_update_text` | Update text in a placeholder on a slide |
| `pptx_update_slide_data` | Update a named or indexed shape while preserving formatting |
| `pptx_insert_image` | Embed an image (PNG, JPG, GIF) on a slide |
| `pptx_extract_talking_points` | Extract the highest-signal talking points from each slide |
| `pptx_export_markdown` | Export a full presentation to a structured markdown file |

→ Full parameter docs and examples: [docs/TOOL_REFERENCE.md](docs/TOOL_REFERENCE.md)

---

## Use Cases

**Meeting Prep Assistant**
Ask your AI assistant to read your keynote deck, pull out the key talking points per slide, and give you a quick summary before you walk into the room. Use `pptx_extract_talking_points` to get ranked, noise-filtered bullet points from every slide in a single call.

**Documentation Generator**
Export a training presentation to markdown for your knowledge base. Use `pptx_export_markdown` to convert the whole deck—headings, bullets, tables, and image references—to a structured `.md` file in one step. Keep the deck as the source of truth; let the agent keep the docs in sync.

**Data Dashboard Updater**
Connect pptx-mcp with a data source MCP. Your agent fetches today's KPIs and updates the metrics slide automatically—no manual editing needed. See the [multi-source composition guide](docs/MULTI_SOURCE_COMPOSITION.md) and the included [mock-data-mcp example](examples/mock-data-mcp/) to run it locally.

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