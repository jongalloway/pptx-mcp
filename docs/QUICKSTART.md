# pptx-mcp Quickstart

Get from zero to a working PowerPoint MCP server in 5–10 minutes.

---

## Prerequisites

Before you begin, make sure you have:

- **[.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)** — verify with `dotnet --version` (must be `10.x.x` or later)
- **[Git](https://git-scm.com/downloads)** — for cloning the repository
- A **PowerPoint file** (`.pptx`) to test with — any `.pptx` file on your machine works
- An **MCP-compatible AI client**, such as:
  - [Claude Desktop](https://claude.ai/download)
  - [VS Code with GitHub Copilot](https://code.visualstudio.com/) (agent mode)
  - Any client that supports the [Model Context Protocol](https://modelcontextprotocol.io)

---

## Installation

### Option A: From Source (available now)

```bash
# 1. Clone the repository
git clone https://github.com/jongalloway/pptx-mcp.git
cd pptx-mcp

# 2. Build the project
dotnet build src/PptxMcp/PptxMcp.csproj --configuration Release

# 3. Note the path to the output binary — you'll need it when configuring your MCP client
#    Default output path: src/PptxMcp/bin/Release/net10.0/PptxMcp
```

Verify the build succeeded:

```bash
dotnet run --project src/PptxMcp/PptxMcp.csproj --configuration Release
# The server starts and waits for MCP messages on stdin — press Ctrl+C to stop
```

### Option B: From NuGet (coming soon)

> **Note:** NuGet package publishing is planned for a future release. Once available, installation will be:
>
> ```bash
> dotnet tool install --global PptxMcp
> ```
>
> This page will be updated with full instructions when the package is published.

---

## Configure Your MCP Client

### Claude Desktop

Open (or create) your Claude Desktop configuration file:

| Platform | Path |
|----------|------|
| macOS    | `~/Library/Application Support/Claude/claude_desktop_config.json` |
| Windows  | `%APPDATA%\Claude\claude_desktop_config.json` |
| Linux    | `~/.config/Claude/claude_desktop_config.json` |

Add the `pptx-mcp` server entry. Replace `/path/to/pptx-mcp` with the actual path where you cloned the repository:

```json
{
  "mcpServers": {
    "pptx-mcp": {
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "/path/to/pptx-mcp/src/PptxMcp/PptxMcp.csproj",
        "--configuration",
        "Release"
      ]
    }
  }
}
```

> **Tip:** For faster startup, build the project first and point directly at the compiled binary:
>
> ```json
> {
>   "mcpServers": {
>     "pptx-mcp": {
>       "command": "/path/to/pptx-mcp/src/PptxMcp/bin/Release/net10.0/PptxMcp"
>     }
>   }
> }
> ```
>
> On Windows, add `.exe` to the binary name: `PptxMcp.exe`.

After saving the file, **restart Claude Desktop** to pick up the new server.

### Other MCP Clients

pptx-mcp uses the standard **stdio transport** defined by the [Model Context Protocol specification](https://modelcontextprotocol.io/docs/concepts/transports). Any MCP client that supports stdio transport can connect to it using the same `command`/`args` pattern shown above. Refer to your client's documentation for the exact configuration syntax.

---

## Run Your First Command

### 1. Open a chat with your MCP client

Once Claude Desktop (or your client) has restarted and connected to the server, you should see `pptx-mcp` listed in the available tools.

### 2. Ask the agent to list your slides

Type a prompt like:

> "List all the slides in `/Users/me/Documents/my-presentation.pptx`"

The agent will call `pptx_list_slides` and return structured output similar to:

```json
[
  {
    "Index": 0,
    "Title": "Introduction",
    "LayoutName": "Title Slide",
    "ShapeCount": 2
  },
  {
    "Index": 1,
    "Title": "Agenda",
    "LayoutName": "Title and Content",
    "ShapeCount": 3
  },
  {
    "Index": 2,
    "Title": "Key Takeaways",
    "LayoutName": "Title and Content",
    "ShapeCount": 4
  }
]
```

### 3. What just happened?

The agent translated your natural-language request into a call to the `pptx_list_slides` MCP tool. pptx-mcp opened the `.pptx` file using the OpenXML SDK, extracted slide metadata (index, title, layout, shape count), and returned it as structured JSON. The agent then presented that data to you in readable form.

### 4. Try a few more commands

```
"Show me the content of slide 0 in my-presentation.pptx"
→ calls pptx_get_slide_content

"What layouts are available in my-presentation.pptx?"
→ calls pptx_list_layouts

"Add a new slide using the 'Title and Content' layout to my-presentation.pptx"
→ calls pptx_add_slide
```

---

## Troubleshooting

### Server doesn't appear in Claude Desktop

- Make sure Claude Desktop was fully **restarted** after editing `claude_desktop_config.json`.
- Verify the path in your config is **absolute**, not relative.
- Run the server manually to check for errors:
  ```bash
  dotnet run --project /path/to/pptx-mcp/src/PptxMcp/PptxMcp.csproj
  ```
- Check the Claude Desktop logs (macOS: `~/Library/Logs/Claude/`) for connection errors.

### `dotnet: command not found`

The .NET 10 SDK is not installed or not on your `PATH`. Download it from [dotnet.microsoft.com](https://dotnet.microsoft.com/download/dotnet/10.0) and follow the installation instructions for your OS.

### `Error: File not found`

The tool received a path that doesn't exist. Make sure you provide an **absolute path** to the `.pptx` file:

```
# ✗ Relative path — may not resolve correctly
"my-presentation.pptx"

# ✓ Absolute path
"/Users/me/Documents/my-presentation.pptx"
"C:\Users\Me\Documents\my-presentation.pptx"
```

### Build fails: SDK version mismatch

The project requires **.NET 10**. If you see a version error, check your installed SDK:

```bash
dotnet --list-sdks
```

Install .NET 10 from [dotnet.microsoft.com/download/dotnet/10.0](https://dotnet.microsoft.com/download/dotnet/10.0) and try again. `global.json` in the repo root pins the SDK version automatically once it is installed.

### Agent says tools are unavailable after connecting

Some clients require you to **explicitly enable** MCP tools or grant permissions. Check your client's settings for an "Allow MCP tools" or "Approve tools" option. In Claude Desktop, tools are enabled by default once the server entry is in the config file.

---

## Next Steps

- **[TOOL_REFERENCE.md](TOOL_REFERENCE.md)** — Complete reference for all 7 available tools with parameter descriptions and example outputs.
- **[docs/EXAMPLES/](EXAMPLES/)** — Worked examples of common agentic workflows: content extraction, slide updates, image insertion, and more.
- **[PRD.md](PRD.md)** — Product roadmap including planned Phase 1 (content extraction) and Phase 2 (data-driven updates) features.
- **[GitHub Issues](https://github.com/jongalloway/pptx-mcp/issues)** — Report bugs, request features, or follow development progress.
