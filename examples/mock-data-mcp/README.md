# mock-data-mcp

A sample MCP server providing mock business metrics and blog post data. Built with the same .NET MCP SDK pattern as pptx-tools. Use it with pptx-tools to run the [multi-source composition example](../../docs/MULTI_SOURCE_COMPOSITION.md) locally without any external API keys or services.

---

## Tools

| Tool | What it does |
|------|-------------|
| `get_weekly_metrics` | Returns weekly KPIs: ARR, MRR, NRR, new logos, churn rate, and highlights |
| `get_team_updates` | Returns department-level status updates (Engineering, Sales, Support, Marketing) |
| `get_latest_blog_posts` | Returns recent blog post summaries, optionally filtered by tag |

---

## Quick Start

**Prerequisites:** [.NET 10 SDK](https://dotnet.microsoft.com/download)

```bash
# From the repo root
dotnet run --project examples/mock-data-mcp --configuration Release
```

The server starts on stdio and waits for MCP messages. It is designed to be registered in your AI client config alongside pptx-tools, not run manually.

---

## Wire Up with pptx-tools (Claude Desktop)

Add both servers to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "pptx-tools": {
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "/absolute/path/to/pptx-tools/src/PptxTools",
        "--configuration",
        "Release"
      ]
    },
    "mock-data-mcp": {
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "/absolute/path/to/pptx-tools/examples/mock-data-mcp",
        "--configuration",
        "Release"
      ]
    }
  }
}
```

Restart Claude Desktop after saving. Both servers will appear in the tool list.

---

## Tool Details

### `get_weekly_metrics`

Returns mock KPIs for a given ISO week. Data varies deterministically by week so different week inputs produce different—but internally consistent—values.

**Parameters**

| Name | Type | Description |
|------|------|-------------|
| `week` | string (optional) | ISO week identifier, e.g. `"2025-W24"`. Defaults to current week. |

**Example response**

```json
{
  "week": "2025-W24",
  "period": "Week of Jun 9–15, 2025",
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
    "NRR reached 115% — second consecutive month above 109%",
    "Support ticket volume down 18% following documentation refresh",
    "New integration: GitHub connector released"
  ],
  "last_updated": "2025-06-09T12:00:00Z"
}
```

---

### `get_team_updates`

Returns department-level status updates. Useful for populating a multi-section "Team Updates" slide.

**Parameters**

| Name | Type | Description |
|------|------|-------------|
| `week` | string (optional) | ISO week identifier. Defaults to current week. |

**Example response (abbreviated)**

```json
{
  "week": "2025-W24",
  "departments": [
    {
      "name": "Engineering",
      "status": "on-track",
      "updates": [
        "Shipped v3.4.0 with 6 bug fixes",
        "Performance: p99 API latency down 12% vs last week",
        "Test coverage at 94% — up 2pp this sprint"
      ]
    },
    {
      "name": "Sales",
      "status": "ahead",
      "updates": [
        "Pipeline: $2100K in late-stage deals",
        "Win rate: 46% (vs 38% 30-day avg)",
        "3 pilots converting to paid this week"
      ]
    }
  ]
}
```

---

### `get_latest_blog_posts`

Returns recent blog post summaries. Filter by tag to focus on specific topics.

**Parameters**

| Name | Type | Description |
|------|------|-------------|
| `tag` | string (optional) | Filter by tag: `"mcp"`, `"dotnet"`, `"ai-agents"`, `"openxml"`, `"documentation"`. |
| `count` | integer (optional) | Max posts to return. Defaults to `5`. |

**Example response (abbreviated)**

```json
{
  "fetched_at": "2025-06-09T12:00:00Z",
  "filter_tag": "mcp",
  "posts": [
    {
      "Published": "2025-06-09",
      "Title": "Introducing MCP Composition Patterns for .NET Agents",
      "Url": "https://devblogs.microsoft.com/dotnet/mcp-composition-patterns",
      "Summary": "Learn how to compose multiple MCP servers...",
      "Tags": ["mcp", "dotnet", "ai-agents", "composition"]
    }
  ]
}
```

---

## Extending This Example

This server is intentionally minimal. To adapt it to real data:

1. Replace the static data in `Tools/MetricsTools.cs` and `Tools/BlogTools.cs` with live API calls (HTTP clients, database queries, etc.).
2. Add authentication as needed in `Program.cs`.
3. Adjust tool names and parameters to match your data model.

The composition pattern with pptx-tools stays the same regardless of how the data source MCP is implemented.

---

## Related

- [Multi-Source Composition Guide](../../docs/MULTI_SOURCE_COMPOSITION.md) — full walkthrough
- [pptx-tools README](../../README.md) — PowerPoint tool reference
- [docs/EXAMPLES.md](../../docs/EXAMPLES.md) — all usage examples
