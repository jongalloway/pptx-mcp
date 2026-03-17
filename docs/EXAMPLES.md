# pptx-mcp Examples

Real-world use case walkthroughs showing how to use pptx-mcp with AI agents.

---

## Table of Contents

1. [Meeting Prep Assistant](#1-meeting-prep-assistant)
2. [Documentation Generator](#2-documentation-generator)
3. [Research Synthesis Tool](#3-research-synthesis-tool)
4. [Data Dashboard Updater](#4-data-dashboard-updater)
5. [Sprint Status Board Builder](#5-sprint-status-board-builder)
6. [Quarterly Metrics Table Updater](#6-quarterly-metrics-table-updater)
7. [Product Feature Comparison Matrix](#7-product-feature-comparison-matrix)

---

## 1. Meeting Prep Assistant

### Scenario

You have a keynote or product deck and need to prepare for a presentation quickly. Instead of skimming dozens of slides manually, you ask an AI agent to extract the key talking points slide-by-slide so you can review and rehearse in minutes.

### Agent Prompt

```
I have a presentation at /presentations/q2-product-review.pptx.
Extract the key talking points from each slide and give me a concise summary organized by slide title.
```

### Tool Workflow

1. **`pptx_extract_talking_points`** — Extract the top ranked talking points from every slide in a single call. The tool filters out noise (formatting-only text, presenter notes labels) and returns bullet-level content ranked by signal.

   ```json
   {
     "name": "pptx_extract_talking_points",
     "arguments": {
       "filePath": "/presentations/q2-product-review.pptx",
       "topN": 5
     }
   }
   ```

2. **Agent synthesizes** — The AI agent formats the returned `Points` arrays per slide into a readable briefing, incorporating speaker notes or slide titles as section headers.

### Example Output

`pptx_extract_talking_points` returns:

```json
[
  {
    "SlideIndex": 0,
    "Title": "Q2 Product Review",
    "Points": []
  },
  {
    "SlideIndex": 1,
    "Title": "Revenue Highlights",
    "Points": [
      "Q2 ARR up 18% YoY",
      "EMEA region grew 34%",
      "Net Revenue Retention: 112%"
    ]
  },
  {
    "SlideIndex": 2,
    "Title": "Roadmap Preview",
    "Points": [
      "GA release: Q3 2025",
      "New integrations: Slack, Teams, Notion",
      "Mobile app entering beta",
      "Pricing page refresh"
    ]
  }
]
```

Agent-synthesized summary:

```
## Slide 1: Revenue Highlights
- Q2 ARR up 18% year-over-year
- EMEA region led growth at 34%
- Net Revenue Retention strong at 112%

## Slide 2: Roadmap Preview
- GA release: Q3 2025
- New integrations: Slack, Teams, Notion
- Mobile app entering beta
- Pricing page refresh
```

### Try It Yourself

1. Open your AI assistant (Claude, Copilot, etc.) with pptx-mcp configured.
2. Point it at any `.pptx` file on your machine.
3. Use the agent prompt above, replacing the file path.
4. Don't have a deck handy? Use any of the [sample files in the test suite](../tests/PptxMcp.Tests/).

---

## 2. Documentation Generator

### Scenario

Your team maintains internal training or onboarding decks as the source of truth, but your documentation site needs markdown. Instead of manually copying slide content, an AI agent exports the full presentation to a structured markdown file that can be checked into your docs repo.

### Agent Prompt

```
Export the presentation at /presentations/onboarding-engineering.pptx to markdown.
Save the result to /docs/onboarding-engineering.md.
```

### Tool Workflow

1. **`pptx_export_markdown`** — Convert the entire presentation to markdown in a single call. Slide titles become headings, body text becomes bullet lists, tables are converted to markdown format, and embedded images are saved to a sibling `{name}_images/` directory with relative references.

   ```json
   {
     "name": "pptx_export_markdown",
     "arguments": {
       "filePath": "/presentations/onboarding-engineering.pptx",
       "outputPath": "/docs/onboarding-engineering.md"
     }
   }
   ```

### Example Output

`pptx_export_markdown` returns the markdown string and writes it to the output file:

```markdown
# Engineering Onboarding

---
<!-- Slide 0 -->

## Welcome to the Team

Welcome to the engineering team. This guide walks you through your first week
setup and key processes.

---
<!-- Slide 1 -->

## Development Environment Setup

- Install .NET 10 SDK
- Clone the repository: `git clone https://github.com/org/repo`
- Run `dotnet build` to verify setup
- Run `dotnet test` to confirm all tests pass

---
<!-- Slide 2 -->

## Code Review Process

| Step | Owner | SLA |
|------|-------|-----|
| Open PR | Author | — |
| Review assigned | Tech lead | 1 business day |
| Approval + merge | Reviewer | 2 business days |

---
<!-- Slide 3 -->

## Team Resources

![team-org-chart](onboarding-engineering_images/slide3_image1.png)

- Org chart and reporting structure above
- Internal wiki: https://wiki.example.com
- Slack: #engineering-onboarding
```

### Try It Yourself

1. Configure pptx-mcp in your AI assistant.
2. Use a training or documentation deck you already have.
3. Ask the agent to export it to markdown using the prompt above.
4. Review the output and check it into your docs repo.

---

## 3. Research Synthesis Tool

### Scenario

You have collected multiple research or competitive analysis decks and need a unified summary of findings. Rather than reading each deck manually and synthesizing by hand, you ask an AI agent to read all presentations, extract key findings from each, and produce a consolidated research brief.

### Agent Prompt

```
I have three research decks in /research/:
- competitor-a-analysis.pptx
- competitor-b-analysis.pptx
- market-trends-2025.pptx

Read all three presentations. For each deck, extract the key findings and
conclusions (focus on slides with titles containing "Finding", "Conclusion",
"Summary", or "Recommendation"). Then produce a consolidated research brief
that compares the three sources by theme.
```

### Tool Workflow

1. **`pptx_list_slides`** (once per file) — Survey slide titles across all three decks to identify relevant slides.

   ```json
   { "filePath": "/research/competitor-a-analysis.pptx" }
   { "filePath": "/research/competitor-b-analysis.pptx" }
   { "filePath": "/research/market-trends-2025.pptx" }
   ```

2. **`pptx_get_slide_content`** (targeted) — Extract content from the slides identified in step 1 as relevant (findings, conclusions, summaries).

   ```json
   { "filePath": "/research/competitor-a-analysis.pptx", "slideIndex": 4 }
   { "filePath": "/research/competitor-a-analysis.pptx", "slideIndex": 9 }
   { "filePath": "/research/competitor-b-analysis.pptx", "slideIndex": 3 }
   // ...
   ```

3. **Agent synthesizes** — The agent groups findings by theme across all three decks and produces a cross-source research brief.

### Example Output

`pptx_list_slides` for `competitor-a-analysis.pptx`:

```json
[
  { "Index": 0, "Title": "Competitor A — Strategic Analysis", "PlaceholderCount": 2 },
  { "Index": 1, "Title": "Market Position", "PlaceholderCount": 3 },
  { "Index": 2, "Title": "Product Comparison", "PlaceholderCount": 4 },
  { "Index": 3, "Title": "Pricing Strategy", "PlaceholderCount": 3 },
  { "Index": 4, "Title": "Key Findings", "PlaceholderCount": 3 },
  { "Index": 5, "Title": "Recommendations", "PlaceholderCount": 3 }
]
```

`pptx_get_slide_content` for slide 4 ("Key Findings"):

```json
{
  "SlideIndex": 4,
  "Shapes": [
    {
      "ShapeType": "Text",
      "PlaceholderType": "title",
      "Text": "Key Findings"
    },
    {
      "ShapeType": "Text",
      "PlaceholderType": "body",
      "Paragraphs": [
        "Competitor A holds 23% market share in enterprise segment",
        "Pricing 15% below market average; margin pressure evident",
        "No MCP or AI integration in current product roadmap",
        "Customer satisfaction scores declining (NPS: 31, down from 44)"
      ]
    }
  ]
}
```

Agent-synthesized research brief:

```markdown
# Consolidated Research Brief — Q2 2025

## Theme: Market Position
- **Competitor A**: 23% enterprise share; pricing 15% below market average
- **Competitor B**: Growing SMB focus; 41% YoY growth in that segment
- **Market Trends 2025**: Enterprise AI adoption accelerating; MCP adoption
  cited as key differentiator in 3 of 5 analyst reports

## Theme: AI & Integration Readiness
- **Competitor A**: No AI integration planned; gap vs. market trend
- **Competitor B**: GPT-4 integration in beta; limited MCP support
- **Market Trends 2025**: 67% of enterprises plan AI-assisted workflow
  tools by end of 2025

## Theme: Customer Satisfaction
...

## Recommendations
1. Prioritize MCP/AI integration to widen competitive gap
2. Monitor Competitor B's SMB push — potential channel conflict
3. Revisit pricing strategy for enterprise tier given Competitor A pressure
```

### Try It Yourself

1. Gather 2–3 related `.pptx` files (research, competitive analysis, reports).
2. Use the agent prompt above, updating the file paths.
3. The agent will make multiple `pptx_list_slides` calls followed by targeted `pptx_get_slide_content` calls on the relevant slides.
4. Ask the agent to save the output to a markdown file for sharing.

---

## 4. Data Dashboard Updater

### Scenario

Your team has a weekly board presentation with a metrics slide. Instead of manually updating KPI values each Monday morning, an AI agent fetches the latest numbers from a data source MCP and updates the relevant slides automatically.

This example uses the **[mock-data-mcp](../examples/mock-data-mcp/)** server included in this repo, which you can run locally without any API keys. The same pattern works with any MCP server that exposes live data.

### Prerequisites

Build both servers from the repo root:

```bash
dotnet build PptxMcp.slnx --configuration Release
dotnet build examples/mock-data-mcp/MockDataMcp.csproj --configuration Release
```

Configure both in your AI client — see [docs/MULTI_SOURCE_COMPOSITION.md](MULTI_SOURCE_COMPOSITION.md) for full setup instructions.

### Agent Prompt

```
I have a weekly board presentation at /presentations/weekly-board.pptx.

1. Fetch this week's KPIs using get_weekly_metrics.
2. Fetch team updates using get_team_updates.
3. List all slides in the presentation to find the right slides to update.
4. Read the KPI Summary slide content to identify placeholder indices.
5. Update the KPI placeholders with the new values: ARR, MRR, NRR, new logos, churn rate.
6. Update the Team Updates slide with the department statuses.
7. Stamp the title slide subtitle with today's date.
```

### Tool Workflow

1. **`get_weekly_metrics`** (mock-data-mcp) — Fetch KPIs: ARR, MRR, NRR, new logos, churn rate, and notable highlights.

   ```json
   { "name": "get_weekly_metrics", "arguments": {} }
   ```

2. **`get_team_updates`** (mock-data-mcp) — Fetch department-level status updates.

   ```json
   { "name": "get_team_updates", "arguments": {} }
   ```

3. **`pptx_list_slides`** (pptx-mcp) — Identify slide titles and indices.

   ```json
   { "name": "pptx_list_slides", "arguments": { "filePath": "/presentations/weekly-board.pptx" } }
   ```

4. **`pptx_get_slide_content`** (pptx-mcp) — Read the KPI slide to map placeholder names to indices before writing.

   ```json
   { "name": "pptx_get_slide_content", "arguments": { "filePath": "/presentations/weekly-board.pptx", "slideIndex": 1 } }
   ```

5. **`pptx_update_slide_data`** (pptx-mcp, repeated) — Write each new KPI value to its named shape, preserving formatting.

   ```json
   { "name": "pptx_update_slide_data", "arguments": { "filePath": "/presentations/weekly-board.pptx", "slideNumber": 2, "shapeName": "ARR", "newText": "$19.6M ARR (+3.3%)" } }
   ```

### Example Output

`get_weekly_metrics` returns:

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
  ]
}
```

After all updates, the agent confirms:

```
✓ Fetched KPIs for 2025-W24
✓ Fetched team updates for 2025-W24
✓ Found 4 slides in weekly-board.pptx
✓ Mapped 5 KPI placeholders on slide 1 (KPI Summary)
✓ Updated ARR → $19.6M ARR (+3.3%)
✓ Updated MRR → $1.63M MRR
✓ Updated NRR → 115% NRR
✓ Updated New Logos → 8 new logos
✓ Updated Churn → 1.7% churn
✓ Updated slide 2 (Team Updates)
✓ Stamped slide 0 with "Week of June 9–15, 2025"

weekly-board.pptx is ready for Monday's board meeting.
```

### Try It Yourself

1. Build both servers (see Prerequisites above).
2. Add both to your AI client config (see [MULTI_SOURCE_COMPOSITION.md](MULTI_SOURCE_COMPOSITION.md)).
3. Point the agent at any `.pptx` file with text placeholders and use the prompt above.
4. For a complete walkthrough with architecture diagrams, two full scenarios, and design guidance, see [docs/MULTI_SOURCE_COMPOSITION.md](MULTI_SOURCE_COMPOSITION.md).

---

- [README](../README.md) — Full tool reference and configuration
- [PRD](PRD.md) — Product requirements, goals, and roadmap
- [Multi-Source Composition Guide](MULTI_SOURCE_COMPOSITION.md) — Architecture, configuration, and full walkthroughs for composing pptx-mcp with external data MCPs
- [mock-data-mcp](../examples/mock-data-mcp/README.md) — Runnable example MCP server providing mock business metrics and blog data

---

## 5. Sprint Status Board Builder

### Scenario

Your engineering team holds a weekly sprint review and needs a slide showing task status at a glance. Instead of manually typing a table into PowerPoint, you ask an AI agent to build the sprint tracking slide automatically from your task list and then patch individual cell values as status changes during the meeting.

### Agent Prompt

```
I have a sprint review deck at /presentations/sprint-42-review.pptx.
Slide 3 is the "Sprint Status" slide. Build a task-status table there
from the following data, then mark the "API Integration" task as "Done"
and the "Load Testing" task as "Blocked – infra outage".
```

### Tool Workflow

1. **`pptx_insert_table`** — Create the sprint status table on slide 3 with header row and one data row per task.

   ```json
   {
     "name": "pptx_insert_table",
     "arguments": {
       "filePath": "/presentations/sprint-42-review.pptx",
       "slideNumber": 3,
       "tableName": "Sprint Status",
       "headers": ["Task", "Owner", "Points", "Status"],
       "rows": [
         ["API Integration",   "Priya",  "5",  "In Progress"],
         ["Auth Refactor",     "Marcus", "8",  "Done"],
         ["Dashboard UI",      "Sofia",  "3",  "In Progress"],
         ["Load Testing",      "Kai",    "5",  "In Progress"],
         ["Docs Update",       "Elena",  "2",  "Done"]
       ],
       "x": 457200,
       "y": 1143000,
       "width": 8229600,
       "height": 2286000
     }
   }
   ```

2. **`pptx_update_table`** — Patch just the two status cells that changed during the meeting.

   ```json
   {
     "name": "pptx_update_table",
     "arguments": {
       "filePath": "/presentations/sprint-42-review.pptx",
       "slideNumber": 3,
       "tableName": "Sprint Status",
       "updates": [
         { "Row": 1, "Column": 3, "Value": "Done" },
         { "Row": 4, "Column": 3, "Value": "Blocked – infra outage" }
       ]
     }
   }
   ```

### Example Output

`pptx_insert_table` returns:

```json
{
  "Success": true,
  "SlideNumber": 3,
  "TableName": "Sprint Status",
  "TableShapeId": 7,
  "TableIndex": 0,
  "RowCount": 6,
  "ColumnCount": 4,
  "Message": "Table inserted successfully."
}
```

`pptx_update_table` returns:

```json
{
  "Success": true,
  "SlideNumber": 3,
  "TableName": "Sprint Status",
  "MatchedBy": "tableName",
  "CellsUpdated": 2,
  "CellsSkipped": 0,
  "Message": "Table updated successfully."
}
```

### Try It Yourself

1. Open any `.pptx` file and choose a blank slide.
2. Call `pptx_insert_table` with your own task list.
3. As statuses change, call `pptx_update_table` with only the cells that need updating — the rest of the table is untouched.

---

## 6. Quarterly Metrics Table Updater

### Scenario

Your finance team maintains a quarterly board deck with a KPI summary table on the executive slide. Each quarter the numbers change but the table structure stays the same. An AI agent reads the current table to confirm the layout, then writes the new values into the exact cells — without touching formatting or other slide content.

### Agent Prompt

```
Open /presentations/q3-board-deck.pptx. Slide 2 has a metrics table named
"KPI Summary". Read the current table to understand its layout, then update
it with Q3 actuals: Revenue $24.1M, Gross Margin 68%, ARR $31.4M,
NRR 114%, New Logos 22, Churn 1.4%.
```

### Tool Workflow

1. **`pptx_get_slide_content`** — Read slide 2 to inspect the table structure before writing.

   ```json
   {
     "name": "pptx_get_slide_content",
     "arguments": {
       "filePath": "/presentations/q3-board-deck.pptx",
       "slideIndex": 1
     }
   }
   ```

2. **`pptx_update_table`** — Write the Q3 actuals into the value column (column 1) of each metric row.

   ```json
   {
     "name": "pptx_update_table",
     "arguments": {
       "filePath": "/presentations/q3-board-deck.pptx",
       "slideNumber": 2,
       "tableName": "KPI Summary",
       "updates": [
         { "Row": 1, "Column": 1, "Value": "$24.1M" },
         { "Row": 2, "Column": 1, "Value": "68%" },
         { "Row": 3, "Column": 1, "Value": "$31.4M" },
         { "Row": 4, "Column": 1, "Value": "114%" },
         { "Row": 5, "Column": 1, "Value": "22" },
         { "Row": 6, "Column": 1, "Value": "1.4%" }
       ]
     }
   }
   ```

### Example Output

`pptx_get_slide_content` for slide 2 shows the existing table:

```json
{
  "SlideIndex": 1,
  "Shapes": [
    {
      "ShapeType": "Text",
      "Name": "Title 1",
      "Text": "Q3 Executive Summary"
    },
    {
      "ShapeType": "Table",
      "Name": "KPI Summary",
      "TableRows": [
        ["Metric",        "Q3 Actual"],
        ["Revenue",       "$22.8M"],
        ["Gross Margin",  "65%"],
        ["ARR",           "$29.7M"],
        ["NRR",           "109%"],
        ["New Logos",     "18"],
        ["Churn",         "1.9%"]
      ]
    }
  ]
}
```

`pptx_update_table` returns:

```json
{
  "Success": true,
  "SlideNumber": 2,
  "TableName": "KPI Summary",
  "MatchedBy": "tableName",
  "CellsUpdated": 6,
  "CellsSkipped": 0,
  "Message": "Table updated successfully."
}
```

### Try It Yourself

1. Open any `.pptx` that already contains a table.
2. Call `pptx_get_slide_content` on the slide to discover table name and row/column layout.
3. Call `pptx_update_table` using `tableName` to target the table precisely and pass just the cells that changed.

---

## 7. Product Feature Comparison Matrix

### Scenario

Your sales team needs a fresh competitive comparison slide for every major prospect meeting. The products and features stay mostly the same, but availability and notes change per prospect. An AI agent builds the full table in one call, then applies per-prospect customizations with a follow-up update.

### Agent Prompt

```
Create a feature comparison table on slide 4 of
/presentations/acme-prospect-deck.pptx. Include our product and two
competitors across five key features. Then mark "SSO / SAML" as
"Coming Q4" for Competitor B since we know Acme cares about that gap.
```

### Tool Workflow

1. **`pptx_insert_table`** — Build the full comparison matrix with headers and all feature rows.

   ```json
   {
     "name": "pptx_insert_table",
     "arguments": {
       "filePath": "/presentations/acme-prospect-deck.pptx",
       "slideNumber": 4,
       "tableName": "Feature Comparison",
       "headers": ["Feature", "Our Product", "Competitor A", "Competitor B"],
       "rows": [
         ["MCP / AI Integration", "✓ Native",   "✗ None",       "Beta"],
         ["SSO / SAML",           "✓ Included", "✓ Add-on (Enterprise tier)", "✗ Roadmap"],
         ["Audit Logs",           "✓ 12 months","✓ 3 months",   "✓ 6 months"],
         ["Mobile App",           "✓ iOS + Android", "iOS only","✗ None"],
         ["SLA",                  "99.99%",     "99.9%",        "99.9%"]
       ],
       "x": 457200,
       "y": 1143000,
       "width": 8229600,
       "height": 2743200
     }
   }
   ```

2. **`pptx_update_table`** — Apply the prospect-specific customization to the SSO row.

   ```json
   {
     "name": "pptx_update_table",
     "arguments": {
       "filePath": "/presentations/acme-prospect-deck.pptx",
       "slideNumber": 4,
       "tableName": "Feature Comparison",
       "updates": [
         { "Row": 2, "Column": 3, "Value": "Coming Q4" }
       ]
     }
   }
   ```

### Example Output

`pptx_insert_table` returns:

```json
{
  "Success": true,
  "SlideNumber": 4,
  "TableName": "Feature Comparison",
  "TableShapeId": 9,
  "TableIndex": 0,
  "RowCount": 6,
  "ColumnCount": 4,
  "Message": "Table inserted successfully."
}
```

`pptx_update_table` returns:

```json
{
  "Success": true,
  "SlideNumber": 4,
  "TableName": "Feature Comparison",
  "MatchedBy": "tableName",
  "CellsUpdated": 1,
  "CellsSkipped": 0,
  "Message": "Table updated successfully."
}
```

### Try It Yourself

1. Prepare a prospect deck with an empty or placeholder slide.
2. Call `pptx_insert_table` once to build the base matrix (reuse the same data for all prospects).
3. Call `pptx_update_table` with only the cells that differ per prospect — the rest of the table is inherited from the base.
4. Combine with `pptx_update_slide_data` to also stamp the title slide with the prospect name.

---

## Related Resources

- [README](../README.md) — Full tool reference and configuration
- [PRD](PRD.md) — Product requirements, goals, and roadmap
- [Multi-Source Composition Guide](MULTI_SOURCE_COMPOSITION.md) — Architecture, configuration, and full walkthroughs for composing pptx-mcp with external data MCPs
- [mock-data-mcp](../examples/mock-data-mcp/README.md) — Runnable example MCP server providing mock business metrics and blog data
