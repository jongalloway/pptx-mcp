# Quickstart: pptx-mcp

Get pptx-mcp running with your AI assistant in a few minutes.

---

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- Claude Desktop (or any MCP-compatible client)

---

## Step 1: Clone and Build

```bash
git clone https://github.com/jongalloway/pptx-mcp.git
cd pptx-mcp
dotnet build PptxMcp.slnx --configuration Release
```

---

## Step 2: Configure Claude Desktop

Open (or create) `claude_desktop_config.json`:

- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

Add the following, replacing the path with the absolute path to your clone:

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

> **Tip:** On Windows, use forward slashes or escaped backslashes in the path: `C:/Users/you/pptx-mcp/src/PptxMcp`

---

## Step 3: Restart Claude Desktop

Fully quit and relaunch Claude Desktop. The MCP server starts automatically when you open a conversation.

---

## Step 4: Try It Out

Point at a `.pptx` file and try some natural language prompts:

```
List all the slides in /path/to/my-presentation.pptx
```

```
Show me the content of slide 0 in /path/to/my-presentation.pptx
```

```
Add a new slide using the "Title and Content" layout to /path/to/my-presentation.pptx
```

---

## What's Next

- [TOOL_REFERENCE.md](TOOL_REFERENCE.md) — full list of tools and parameters
- [PRD.md](PRD.md) — roadmap and planned capabilities
