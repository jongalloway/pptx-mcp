# pptx-mcp Client Setup Guide

This guide shows how to configure different MCP clients to use **pptx-mcp**, a .NET-based MCP server for reading and modifying PowerPoint (.pptx) files.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Claude Desktop](#claude-desktop)
- [VS Code Extensions](#vs-code-extensions)
  - [GitHub Copilot (VS Code)](#github-copilot-vs-code)
  - [Cline](#cline)
  - [Codeium / Windsurf](#codeium--windsurf)
- [Command-Line / Custom Clients](#command-line--custom-clients)
- [Local LLMs](#local-llms)
- [Configuring Multiple MCPs](#configuring-multiple-mcps)
- [Composition Tips](#composition-tips)
- [Troubleshooting](#troubleshooting)

---

## Prerequisites

- **.NET 10 SDK** installed and on your `PATH`. Verify with:
  ```bash
  dotnet --version
  # Should output 10.x.x
  ```
- The **pptx-mcp** binary available either as a .NET global tool or built from source (see [Installation](#installation) below).

---

## Installation

### Option A — Install as a .NET global tool (recommended)

```bash
dotnet tool install --global PptxMcp
```

Verify the tool is on your `PATH`:

```bash
pptx-mcp --version
```

> **Note:** If `pptx-mcp` is not found after installation, ensure the .NET global tools directory is on your `PATH`:
> - **Windows:** `%USERPROFILE%\.dotnet\tools`
> - **macOS/Linux:** `~/.dotnet/tools`

### Option B — Run from source

Clone the repository and note the path to the project file:

```bash
git clone https://github.com/jongalloway/pptx-mcp
cd pptx-mcp
dotnet build src/PptxMcp/PptxMcp.csproj --configuration Release
```

When configuring clients below, replace `"command": "pptx-mcp"` with:
```json
"command": "dotnet",
"args": ["run", "--project", "/absolute/path/to/pptx-mcp/src/PptxMcp", "--"]
```

---

## Claude Desktop

### Prerequisites

- [Claude Desktop](https://claude.ai/download) installed
- pptx-mcp installed (see [Installation](#installation))

### Configuration

Open your Claude Desktop configuration file:

| Platform | Path |
|----------|------|
| **macOS** | `~/Library/Application Support/Claude/claude_desktop_config.json` |
| **Windows** | `%APPDATA%\Claude\claude_desktop_config.json` |

Add `pptx-mcp` to the `mcpServers` section:

```json
{
  "mcpServers": {
    "pptx-mcp": {
      "command": "pptx-mcp"
    }
  }
}
```

If you are running from source instead of the global tool:

```json
{
  "mcpServers": {
    "pptx-mcp": {
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "/absolute/path/to/pptx-mcp/src/PptxMcp",
        "--"
      ]
    }
  }
}
```

**Restart Claude Desktop** after editing the config file.

### Verify it works

In a Claude Desktop conversation, ask:

> "Use pptx-mcp to list the slides in `/path/to/my/presentation.pptx`"

Claude should invoke `pptx_list_slides` and return the slide list. You can also confirm the server loaded by opening **Claude Desktop → Settings → Developer → MCP Servers** and checking that `pptx-mcp` shows a green status indicator.

### Troubleshooting

| Problem | Fix |
|---------|-----|
| `pptx-mcp` not found | Ensure `~/.dotnet/tools` (macOS/Linux) or `%USERPROFILE%\.dotnet\tools` (Windows) is on your `PATH` |
| Server shows red/error status | Check Claude Desktop logs: `~/Library/Logs/Claude/mcp*.log` (macOS) or `%APPDATA%\Claude\logs\` (Windows) |
| JSON parse error in config | Validate `claude_desktop_config.json` with a JSON linter — trailing commas and comments are not allowed |
| File not found errors from tools | Use **absolute paths** when passing `.pptx` file paths to tools |

---

## VS Code Extensions

### GitHub Copilot (VS Code)

VS Code 1.99+ with the GitHub Copilot extension supports MCP servers natively.

#### Workspace configuration

Create or edit `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "pptx-mcp": {
      "type": "stdio",
      "command": "pptx-mcp"
    }
  }
}
```

#### User-level configuration

To make pptx-mcp available in all workspaces, add it to your VS Code user settings (`settings.json`):

```json
{
  "mcp": {
    "servers": {
      "pptx-mcp": {
        "type": "stdio",
        "command": "pptx-mcp"
      }
    }
  }
}
```

#### Verify it works

1. Open the Copilot Chat panel (`Ctrl+Alt+I` / `⌃⌘I`)
2. Switch to **Agent mode** using the mode selector
3. Ask: `"List the slides in /path/to/my/presentation.pptx using pptx-mcp"`

Copilot will prompt you to approve the tool call, then display results.

---

### Cline

[Cline](https://github.com/cline/cline) is a VS Code extension that supports MCP servers.

#### Configuration

1. Open VS Code Settings (`Ctrl+,` / `⌘,`)
2. Search for **Cline MCP** or navigate to **Extensions → Cline → MCP Servers**
3. Add the following to Cline's MCP server configuration:

```json
{
  "pptx-mcp": {
    "command": "pptx-mcp",
    "args": [],
    "disabled": false,
    "autoApprove": []
  }
}
```

Alternatively, edit Cline's settings file directly:

| Platform | Path |
|----------|------|
| **macOS** | `~/Library/Application Support/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json` |
| **Windows** | `%APPDATA%\Code\User\globalStorage\saoudrizwan.claude-dev\settings\cline_mcp_settings.json` |

#### Verify it works

In Cline's chat, ask:
> `"List slides in /path/to/presentation.pptx"`

Cline will identify the available `pptx_list_slides` tool and invoke it automatically.

---

### Codeium / Windsurf

[Windsurf](https://codeium.com/windsurf) (Codeium's editor) supports MCP servers.

#### Configuration

Edit your Windsurf MCP configuration file:

| Platform | Path |
|----------|------|
| **macOS** | `~/.codeium/windsurf/mcp_config.json` |
| **Windows** | `%USERPROFILE%\.codeium\windsurf\mcp_config.json` |

```json
{
  "mcpServers": {
    "pptx-mcp": {
      "command": "pptx-mcp",
      "args": []
    }
  }
}
```

---

## Command-Line / Custom Clients

pptx-mcp uses **stdio transport**: it reads JSON-RPC messages from `stdin` and writes responses to `stdout`. This makes it easy to integrate with any MCP-compatible client or custom tooling.

### Direct invocation

```bash
pptx-mcp
```

The server starts and waits for MCP messages on `stdin`. Send a well-formed MCP `initialize` request followed by tool calls.

### Example: pipe a raw request

```bash
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"0.0.1"}}}' | pptx-mcp
```

### Using with the MCP CLI inspector

The [MCP Inspector](https://github.com/modelcontextprotocol/inspector) is useful for testing and debugging:

```bash
npx @modelcontextprotocol/inspector pptx-mcp
```

This opens an interactive UI for calling tools, inspecting responses, and verifying your server is behaving correctly.

### Using with Python MCP client SDK

```python
import asyncio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

async def main():
    server_params = StdioServerParameters(
        command="pptx-mcp",
        args=[]
    )
    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            result = await session.call_tool(
                "pptx_list_slides",
                {"filePath": "/path/to/presentation.pptx"}
            )
            print(result)

asyncio.run(main())
```

---

## Local LLMs

Any local LLM framework that supports the Model Context Protocol can use pptx-mcp.

### LM Studio

[LM Studio](https://lmstudio.ai/) 0.3.6+ supports MCP servers.

1. Open **LM Studio → Settings → MCP**
2. Click **Add Server**
3. Set **Type** to `stdio`, **Command** to `pptx-mcp`, leave **Args** empty
4. Save and restart the LM Studio chat session

### Ollama + Open WebUI

[Open WebUI](https://github.com/open-webui/open-webui) 0.4+ supports MCP through its Tools integration. Configure pptx-mcp as a stdio tool server in Open WebUI's admin panel under **Settings → Tools → MCP Servers**.

### llm (Simon Willison's CLI)

Install the [llm-mcp](https://github.com/simonw/llm-mcp) plugin and register pptx-mcp:

```bash
pip install llm llm-mcp
llm mcp add pptx-mcp pptx-mcp
llm -m gpt-4o --mcp pptx-mcp "List the slides in /path/to/presentation.pptx"
```

### Generic stdio configuration

For any framework that accepts stdio MCP servers, the invocation is simply:

```
command: pptx-mcp
args:    (none)
env:     (no special environment variables required)
```

---

## Configuring Multiple MCPs

pptx-mcp is designed to be composed with other MCP servers. A common pattern is to pair it with a data source MCP so an AI agent can fetch live data and update slides in a single prompt — no glue code required.

### Claude Desktop — multiple servers

Add each server as a separate entry under `mcpServers`. Claude Desktop loads all configured servers at startup and makes all their tools available to the agent simultaneously:

```json
{
  "mcpServers": {
    "pptx-mcp": {
      "command": "pptx-mcp"
    },
    "mock-data-mcp": {
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "/absolute/path/to/pptx-mcp/examples/mock-data-mcp",
        "--configuration",
        "Release"
      ]
    }
  }
}
```

### VS Code (Copilot) — multiple servers

Add each server to `.vscode/mcp.json`:

```json
{
  "servers": {
    "pptx-mcp": {
      "type": "stdio",
      "command": "pptx-mcp"
    },
    "mock-data-mcp": {
      "type": "stdio",
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "${workspaceFolder}/examples/mock-data-mcp",
        "--configuration",
        "Release"
      ]
    }
  }
}
```

The agent sees tools from all configured servers and selects the right one based on each tool's description.

See [docs/MULTI_SOURCE_COMPOSITION.md](MULTI_SOURCE_COMPOSITION.md) for full configuration examples and step-by-step walkthroughs, including the built-in [mock-data-mcp](../examples/mock-data-mcp/) server you can run locally without any API keys.

---

## Composition Tips

These guidelines apply when pptx-mcp is used alongside other MCP servers:

- **Inspect before writing.** Call `pptx_get_slide_content` on the target slide before making updates. This gives the agent the exact shape names and indices needed to target the right element, and avoids overwriting the wrong content.

- **Prefer `pptx_update_slide_data` for named shapes.** When a shape has a descriptive name (visible as `Name` in `pptx_get_slide_content`), use `pptx_update_slide_data` with `shapeName`. It preserves the shape's existing formatting and is more resilient to deck layout changes than index-based addressing.

- **Anchor updates to slide titles.** In agent prompts, say "update the slide titled 'KPI Summary'" rather than "update slide 3". Slide positions can shift as decks evolve; titles are stable.

- **Be explicit about which server to use for each step.** Telling the agent "use `get_weekly_metrics` from mock-data-mcp, then update the deck using pptx-mcp" reduces ambiguity and prevents the agent from guessing which tool to invoke.

- **Specify the update scope.** Clarify which placeholders to change and which to leave alone — for example, "update the body placeholder but keep the title unchanged". This prevents unintended overwrites.

- **Use absolute paths.** Always pass absolute file paths to pptx-mcp tools. Relative paths resolve against the server's working directory, which may not be what you expect.

---

## Troubleshooting

### `pptx-mcp` command not found

Ensure the .NET global tools directory is on your `PATH`:

```bash
# macOS/Linux — add to ~/.bashrc, ~/.zshrc, or equivalent
export PATH="$HOME/.dotnet/tools:$PATH"

# Windows PowerShell — add to your $PROFILE
$env:PATH += ";$env:USERPROFILE\.dotnet\tools"
```

Then verify:
```bash
pptx-mcp --version
```

### Server starts but tools are not listed

The MCP handshake may be failing. Test with the inspector:

```bash
npx @modelcontextprotocol/inspector pptx-mcp
```

Look for errors in the **Notifications** panel. Common causes:
- `.NET 10 runtime` not installed — install from [https://dot.net](https://dot.net)
- A newer version of `pptx-mcp` requires a higher .NET SDK version

### File not found errors

Always pass **absolute paths** to tool parameters. Relative paths are resolved relative to the server's working directory (the directory where the client launched `pptx-mcp`), which may not be what you expect.

```
# ❌ Ambiguous
"filePath": "my-deck.pptx"

# ✅ Unambiguous
"filePath": "/Users/alice/Documents/my-deck.pptx"
```

### Permission errors on macOS

macOS Gatekeeper may block the first run. Open **System Settings → Privacy & Security** and allow `pptx-mcp` (or `dotnet`) to run, or use:

```bash
xattr -d com.apple.quarantine "$(which pptx-mcp)"
```

### Logging and diagnostics

pptx-mcp writes diagnostic logs to **stderr**. Most MCP clients capture these separately from tool output. To capture them manually:

```bash
pptx-mcp 2>pptx-mcp-debug.log
```

Review `pptx-mcp-debug.log` for startup errors, unhandled exceptions, or OpenXML parsing warnings.

### The server exits immediately

If the server exits without waiting for input, ensure your client is keeping `stdin` open. The server blocks on `stdin`; if the pipe is closed at the client side immediately, the server will terminate normally.

---

## Available Tools

Once connected, the following MCP tools are available:

| Tool | Description |
|------|-------------|
| `pptx_list_slides` | List all slides with metadata |
| `pptx_list_layouts` | List available slide layouts |
| `pptx_get_slide_content` | Get structured content (shapes, text, positions) for a slide |
| `pptx_get_slide_xml` | Get the raw XML of a slide (advanced) |
| `pptx_add_slide` | Add a new slide with a specified layout |
| `pptx_update_text` | Update the text of a placeholder on a slide by index |
| `pptx_update_slide_data` | Update a named or indexed slide shape while preserving formatting |
| `pptx_insert_image` | Insert an image onto a slide |
| `pptx_extract_talking_points` | Extract ranked talking points from each slide |
| `pptx_export_markdown` | Export a full presentation to a structured markdown file |

A good first test after connecting is to call `pptx_list_slides` with a known `.pptx` file path. A successful response confirms the server is connected and operational.
