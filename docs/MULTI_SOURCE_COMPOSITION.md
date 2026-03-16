# Multi-Source Composition with pptx-mcp

This guide shows how to combine pptx-mcp with one or more external data source MCPs so an AI agent can fetch live information and update a presentation in a single prompt — no manual editing, no glue code.

---

## Table of Contents

1. [The Pattern](#1-the-pattern)
2. [Prerequisites](#2-prerequisites)
3. [Configuration](#3-configuration)
4. [Scenario A — Weekly Board Update (mock-data-mcp)](#4-scenario-a--weekly-board-update)
5. [Scenario B — Blog-to-Deck Pipeline (fetch MCP)](#5-scenario-b--blog-to-deck-pipeline)
6. [Architectural Notes](#6-architectural-notes)

---

## 1. The Pattern

Multi-source composition is the practice of configuring an AI agent with two or more MCP servers and letting it orchestrate them to accomplish a task that neither server could handle alone.

```
┌─────────────────────────────────────────────────────────────────┐
│                        AI Agent                                 │
│                                                                 │
│  1. "What are this week's KPIs?"   2. "Update slide 3 with…"   │
└───────────────┬─────────────────────────────┬───────────────────┘
                │                             │
        ┌───────▼────────┐           ┌────────▼────────┐
        │  Data Source   │           │   pptx-mcp      │
        │     MCP        │           │                 │
        │                │           │ pptx_list_slides│
        │ get_weekly_    │           │ pptx_get_slide_ │
        │ metrics        │           │   content       │
        │ get_team_      │           │ pptx_update_    │
        │ updates        │           │   text          │
        │ get_latest_    │           │ pptx_add_slide  │
        │ blog_posts     │           │ pptx_insert_    │
        └────────────────┘           │   image         │
                                     └─────────────────┘
```

**Key properties of this pattern:**

- **No glue code.** The agent handles the orchestration — reading data from one MCP, understanding the presentation structure, and writing updates via the other.
- **Stateless tools.** Each tool call is self-contained. The agent maintains all context between calls.
- **Composable by design.** Any MCP server that returns data can be paired with pptx-mcp. The pattern scales from mock data to production APIs.

---

## 2. Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- pptx-mcp built from source (see [README](../README.md))
- An AI client that supports multiple MCP servers (Claude Desktop, VS Code Copilot, etc.)

For **Scenario A**, also build the mock data server:

```bash
dotnet build examples/mock-data-mcp/MockDataMcp.csproj --configuration Release
```

For **Scenario B**, install the fetch MCP (Node.js required):

```bash
npm install -g @modelcontextprotocol/server-fetch
```

---

## 3. Configuration

### Claude Desktop

Add both servers to `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) or `%APPDATA%\Claude\claude_desktop_config.json` (Windows):

**Scenario A — pptx-mcp + mock-data-mcp**

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

**Scenario B — pptx-mcp + fetch MCP**

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
    },
    "fetch": {
      "command": "npx",
      "args": ["-y", "@modelcontextprotocol/server-fetch"]
    }
  }
}
```

### VS Code (Copilot)

Add an `.vscode/mcp.json` file at the workspace root:

```json
{
  "servers": {
    "pptx-mcp": {
      "type": "stdio",
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "${workspaceFolder}/src/PptxMcp",
        "--configuration",
        "Release"
      ]
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

---

## 4. Scenario A — Weekly Board Update

**Use case:** Your team runs a weekly board presentation with a live metrics slide. Instead of manually updating KPI values each Monday, an AI agent fetches the latest numbers and updates the deck in one prompt.

### Sample Presentation Structure

The example assumes a deck with this structure (create it yourself or use any `.pptx`):

| Slide | Title | Content |
|-------|-------|---------|
| 0 | Weekly Business Review | Title slide |
| 1 | KPI Summary | ARR, MRR, NRR, New Logos, Churn (text placeholders) |
| 2 | Team Updates | Department status bullets |
| 3 | Highlights | Notable wins / notable risks |

### Agent Prompt

```
I have a weekly board presentation at /presentations/weekly-board.pptx.

1. Fetch this week's business KPIs using get_weekly_metrics.
2. Fetch team updates using get_team_updates.
3. List all slides in the presentation to identify which slides to update.
4. Read the current content of the KPI Summary slide (slide 1) and Team Updates
   slide (slide 2) to find the right placeholders.
5. Update slide 1 with the new KPI values: ARR, MRR, NRR, new logos, and churn rate.
6. Update slide 2 with the department statuses and highlights.
7. Update slide 0's subtitle with today's date.

Keep the existing formatting — only change the text values.
```

### Step-by-Step Tool Call Sequence

**Step 1 — Fetch metrics from mock-data-mcp**

```json
{
  "server": "mock-data-mcp",
  "name": "get_weekly_metrics",
  "arguments": {}
}
```

Response (excerpt):
```json
{
  "week": "2025-W24",
  "kpis": {
    "arr_millions": 19.6,
    "arr_change_pct": 3.3,
    "mrr_thousands": 1633,
    "nrr_pct": 115,
    "new_logos": 8,
    "churn_rate_pct": 1.7
  },
  "highlights": [
    "Closed 5 enterprise deals in EMEA",
    "NRR reached 115% — second consecutive month above 109%"
  ]
}
```

**Step 2 — Fetch team updates**

```json
{
  "server": "mock-data-mcp",
  "name": "get_team_updates",
  "arguments": {}
}
```

**Step 3 — Survey the presentation**

```json
{
  "server": "pptx-mcp",
  "name": "pptx_list_slides",
  "arguments": { "filePath": "/presentations/weekly-board.pptx" }
}
```

Response:
```json
[
  { "Index": 0, "Title": "Weekly Business Review", "PlaceholderCount": 2 },
  { "Index": 1, "Title": "KPI Summary", "PlaceholderCount": 6 },
  { "Index": 2, "Title": "Team Updates", "PlaceholderCount": 3 },
  { "Index": 3, "Title": "Highlights", "PlaceholderCount": 2 }
]
```

**Step 4 — Inspect the KPI slide to identify placeholder indices**

```json
{
  "server": "pptx-mcp",
  "name": "pptx_get_slide_content",
  "arguments": {
    "filePath": "/presentations/weekly-board.pptx",
    "slideIndex": 1
  }
}
```

Response (excerpt):
```json
{
  "SlideIndex": 1,
  "Shapes": [
    { "Name": "Title 1", "PlaceholderType": "title", "PlaceholderIndex": 0, "Text": "KPI Summary" },
    { "Name": "ARR", "PlaceholderType": "body", "PlaceholderIndex": 1, "Text": "$18.4M ARR (+2.8%)" },
    { "Name": "MRR", "PlaceholderType": "body", "PlaceholderIndex": 2, "Text": "$1.53M MRR" },
    { "Name": "NRR", "PlaceholderType": "body", "PlaceholderIndex": 3, "Text": "112% NRR" },
    { "Name": "New Logos", "PlaceholderType": "body", "PlaceholderIndex": 4, "Text": "6 new logos" },
    { "Name": "Churn", "PlaceholderType": "body", "PlaceholderIndex": 5, "Text": "1.9% churn" }
  ]
}
```

**Step 5 — Update KPI placeholders with fresh data**

The agent makes one `pptx_update_text` call per placeholder:

```json
{
  "server": "pptx-mcp",
  "name": "pptx_update_text",
  "arguments": {
    "filePath": "/presentations/weekly-board.pptx",
    "slideIndex": 1,
    "placeholderIndex": 1,
    "text": "$19.6M ARR (+3.3%)"
  }
}
```

```json
{
  "server": "pptx-mcp",
  "name": "pptx_update_text",
  "arguments": {
    "filePath": "/presentations/weekly-board.pptx",
    "slideIndex": 1,
    "placeholderIndex": 2,
    "text": "$1.63M MRR"
  }
}
```

*(… repeated for NRR, New Logos, and Churn)*

**Step 6 — Update team updates slide**

After inspecting slide 2 with `pptx_get_slide_content`, the agent updates the body placeholder:

```json
{
  "server": "pptx-mcp",
  "name": "pptx_update_text",
  "arguments": {
    "filePath": "/presentations/weekly-board.pptx",
    "slideIndex": 2,
    "placeholderIndex": 1,
    "text": "Engineering: on-track — shipped v3.4.0, latency down 12%\nSales: ahead — $2.1M pipeline, 46% win rate\nSupport: on-track — CSAT 4.6/5.0, 2h first response\nMarketing: on-track — 1,400 blog views, 330 webinar registrations"
  }
}
```

**Step 7 — Stamp the date**

```json
{
  "server": "pptx-mcp",
  "name": "pptx_update_text",
  "arguments": {
    "filePath": "/presentations/weekly-board.pptx",
    "slideIndex": 0,
    "placeholderIndex": 1,
    "text": "Week of June 9–15, 2025"
  }
}
```

### Expected Final Output

The agent confirms each update:

```
✓ Fetched KPIs for 2025-W24 from mock-data-mcp
✓ Fetched team updates for 2025-W24 from mock-data-mcp
✓ Listed 4 slides in weekly-board.pptx
✓ Inspected slide 1 (KPI Summary) — found 5 data placeholders
✓ Updated ARR → $19.6M ARR (+3.3%)
✓ Updated MRR → $1.63M MRR
✓ Updated NRR → 115% NRR
✓ Updated New Logos → 8 new logos
✓ Updated Churn → 1.7% churn
✓ Updated slide 2 (Team Updates) with department status
✓ Stamped slide 0 with "Week of June 9–15, 2025"

weekly-board.pptx is ready for the Monday board meeting.
```

---

## 5. Scenario B — Blog-to-Deck Pipeline

**Use case:** Keep a "What's New" slide in a product deck synchronized with the latest blog posts. The agent fetches post titles and summaries via the fetch MCP (or mock-data-mcp's `get_latest_blog_posts`) and rewrites the slide in one prompt.

### Agent Prompt (using fetch MCP)

```
I have a product deck at /presentations/product-updates.pptx.

1. Fetch the latest posts from https://devblogs.microsoft.com/dotnet/feed/
   using the fetch tool. Extract the 3 most recent post titles and one-line
   summaries from the RSS feed.
2. List all slides to find the "What's New" or "Recent Updates" slide.
3. Read the current content of that slide.
4. Rewrite the slide body with the 3 new post titles and summaries as bullet
   points. Keep the title placeholder unchanged.
```

### Agent Prompt (using mock-data-mcp)

```
I have a product deck at /presentations/product-updates.pptx.

1. Call get_latest_blog_posts with tag="dotnet" and count=3 to get recent posts.
2. List all slides to find the "What's New" slide.
3. Read the current content of that slide to find the body placeholder index.
4. Rewrite the slide body with the 3 post titles and summaries as bullet points.
```

### Tool Call Sequence (mock-data-mcp variant)

**Fetch posts:**

```json
{
  "server": "mock-data-mcp",
  "name": "get_latest_blog_posts",
  "arguments": { "tag": "dotnet", "count": 3 }
}
```

Response:
```json
{
  "fetched_at": "2025-06-09T12:00:00Z",
  "filter_tag": "dotnet",
  "posts": [
    {
      "Published": "2025-06-09",
      "Title": "Introducing MCP Composition Patterns for .NET Agents",
      "Summary": "Learn how to compose multiple MCP servers to build powerful multi-source agent workflows."
    },
    {
      "Published": "2025-06-03",
      "Title": "What's New in .NET 10 Preview 4",
      "Summary": "Performance improvements in JIT and GC, new LINQ overloads, and System.AI namespace."
    },
    {
      "Published": "2025-05-28",
      "Title": "Building Intelligent Agents with ModelContextProtocol SDK",
      "Summary": "Deep dive into the MCP SDK: tool registration, stdio transport, and composable tools."
    }
  ]
}
```

**Inspect slide to find the "What's New" slide:**

```json
{
  "server": "pptx-mcp",
  "name": "pptx_list_slides",
  "arguments": { "filePath": "/presentations/product-updates.pptx" }
}
```

**Read and update the body placeholder:**

```json
{
  "server": "pptx-mcp",
  "name": "pptx_update_text",
  "arguments": {
    "filePath": "/presentations/product-updates.pptx",
    "slideIndex": 4,
    "placeholderIndex": 1,
    "text": "Introducing MCP Composition Patterns for .NET Agents — Compose multiple MCP servers for powerful multi-source workflows.\nWhat's New in .NET 10 Preview 4 — JIT/GC performance, new LINQ overloads, System.AI namespace.\nBuilding Intelligent Agents with ModelContextProtocol SDK — Tool registration, stdio transport, composable tools."
  }
}
```

---

## 6. Architectural Notes

### Why this works without glue code

Each MCP tool is a discrete, stateless operation. The AI agent functions as the orchestrator: it holds the data returned by one tool in its context window and passes relevant parts as arguments to the next tool. No connector layer, no event bus, no shared state.

### Scaling beyond two servers

The same pattern extends to three or more servers. For example:

```
Agent
 ├── mock-data-mcp      → business KPIs
 ├── fetch (web MCP)    → blog posts, press releases
 └── pptx-mcp           → slide inspection + updates
```

The agent decides which server to query at each step based on the task. MCP tool descriptions (the `Description` attribute set from XML doc comments in .NET) are the primary signal the agent uses to pick the right tool.

### Replacing mock-data-mcp with a real data source

`mock-data-mcp` is designed as a drop-in demonstration. Replace it with any MCP server that exposes your data:

- A database query MCP (Postgres, SQLite, etc.)
- A REST API wrapper MCP (Salesforce, HubSpot, Datadog, etc.)
- The [GitHub MCP server](https://github.com/github/github-mcp-server) for repo-level metrics
- A custom server following the same .NET pattern used here

The pptx-mcp side of the workflow stays unchanged.

### Using `pptx_update_text` vs. a future `pptx_update_slide_data`

The current examples use `pptx_update_text`, which updates one placeholder at a time. This works but requires:
1. An inspection call (`pptx_get_slide_content`) to find placeholder indices
2. One `pptx_update_text` call per value

A future `pptx_update_slide_data` tool (Phase 2) will accept a data map and update all matching placeholders in a single call, streamlining the pattern for data-heavy slides.

### Agent prompt design tips

- **Be explicit about the source of truth.** Tell the agent which MCP to use for each step ("use `get_weekly_metrics` from mock-data-mcp…").
- **Instruct the agent to inspect before writing.** A `pptx_get_slide_content` call before `pptx_update_text` ensures the agent targets the right placeholder index and avoids overwriting the wrong content.
- **Anchor updates to slide titles.** Saying "update the slide titled 'KPI Summary'" is more robust than "update slide 3" if the deck structure might change.
- **Specify the update scope.** Clarify which placeholders to update and which to leave alone (e.g., "keep the title unchanged").

---

## Related Resources

- [mock-data-mcp README](../examples/mock-data-mcp/README.md) — mock server setup and tool reference
- [docs/EXAMPLES.md](EXAMPLES.md) — all pptx-mcp usage examples
- [docs/TOOL_REFERENCE.md](TOOL_REFERENCE.md) — full pptx-mcp tool reference
- [MCP SDK for .NET](https://github.com/modelcontextprotocol/csharp-sdk) — build your own data source MCP
