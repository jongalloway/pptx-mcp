# pptx-mcp Examples

Real-world use case walkthroughs showing how to use pptx-mcp with AI agents.

---

## Table of Contents

1. [Meeting Prep Assistant](#1-meeting-prep-assistant)
2. [Documentation Generator](#2-documentation-generator)
3. [Research Synthesis Tool](#3-research-synthesis-tool)
4. [Data Dashboard Updater *(Coming in Phase 2)*](#4-data-dashboard-updater-coming-in-phase-2)

---

## 1. Meeting Prep Assistant

### Scenario

You have a keynote or product deck and need to prepare for a presentation quickly. Instead of skimming dozens of slides manually, you ask an AI agent to extract the key talking points slide-by-slide so you can review and rehearse in minutes.

### Agent Prompt

```
I have a presentation at /presentations/q2-product-review.pptx.
List all the slides, then for each slide extract the key talking points.
Give me a concise summary organized by slide title.
```

### Tool Workflow

1. **`pptx_list_slides`** — Get an overview of the deck structure (slide count, titles, notes).

   ```json
   { "filePath": "/presentations/q2-product-review.pptx" }
   ```

2. **`pptx_get_slide_content`** (repeated per slide) — Retrieve structured content for each slide: shapes, placeholder text, bullet paragraphs.

   ```json
   { "filePath": "/presentations/q2-product-review.pptx", "slideIndex": 0 }
   { "filePath": "/presentations/q2-product-review.pptx", "slideIndex": 1 }
   // ... continue for each slide
   ```

3. **Agent synthesizes** — The AI agent reads the `Paragraphs` arrays from each `Text`-type shape, filters out decorative or empty shapes, and produces a talking-points summary.

### Example Output

`pptx_list_slides` returns:

```json
[
  { "Index": 0, "Title": "Q2 Product Review", "Notes": null, "PlaceholderCount": 2 },
  { "Index": 1, "Title": "Revenue Highlights", "Notes": "Mention the EMEA uptick", "PlaceholderCount": 3 },
  { "Index": 2, "Title": "Roadmap Preview", "Notes": null, "PlaceholderCount": 4 }
]
```

`pptx_get_slide_content` for slide 1 returns:

```json
{
  "SlideIndex": 1,
  "SlideWidthEmu": 9144000,
  "SlideHeightEmu": 5143500,
  "Shapes": [
    {
      "ShapeId": 2,
      "Name": "Title 1",
      "ShapeType": "Text",
      "IsPlaceholder": true,
      "PlaceholderType": "title",
      "Text": "Revenue Highlights",
      "Paragraphs": ["Revenue Highlights"]
    },
    {
      "ShapeId": 3,
      "Name": "Content Placeholder 2",
      "ShapeType": "Text",
      "IsPlaceholder": true,
      "PlaceholderType": "body",
      "Text": "Q2 ARR up 18% YoY\nEMEA region grew 34%\nNet Revenue Retention: 112%",
      "Paragraphs": [
        "Q2 ARR up 18% YoY",
        "EMEA region grew 34%",
        "Net Revenue Retention: 112%"
      ]
    }
  ]
}
```

Agent-synthesized summary:

```
## Slide 1: Revenue Highlights
- Q2 ARR up 18% year-over-year
- EMEA region led growth at 34%
- Net Revenue Retention strong at 112%
- (Speaker note: Mention the EMEA uptick)

## Slide 2: Roadmap Preview
...
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
Use slide titles as headings and preserve all bullet points.
```

### Tool Workflow

1. **`pptx_list_slides`** — Enumerate slides and capture titles to use as markdown headings.

   ```json
   { "filePath": "/presentations/onboarding-engineering.pptx" }
   ```

2. **`pptx_get_slide_content`** (repeated per slide) — Extract all text shapes, paragraphs, and tables from each slide.

   ```json
   { "filePath": "/presentations/onboarding-engineering.pptx", "slideIndex": 0 }
   { "filePath": "/presentations/onboarding-engineering.pptx", "slideIndex": 1 }
   // ... continue for each slide
   ```

3. **Agent assembles markdown** — Using the structured `Paragraphs` and `TableRows` data, the agent constructs a markdown document:
   - Slide title → `## Heading`
   - Body bullet paragraphs → `- list items`
   - Table shapes → markdown tables
   - Speaker notes → `> blockquote`

4. **Agent writes the file** — The agent saves the markdown output to the target path using its file writing capability.

### Example Output

Input slide content (from `pptx_get_slide_content`):

```json
{
  "SlideIndex": 2,
  "Shapes": [
    {
      "ShapeType": "Text",
      "PlaceholderType": "title",
      "Text": "Development Environment Setup",
      "Paragraphs": ["Development Environment Setup"]
    },
    {
      "ShapeType": "Text",
      "PlaceholderType": "body",
      "Paragraphs": [
        "Install .NET 10 SDK",
        "Clone the repository: git clone https://github.com/org/repo",
        "Run dotnet build to verify setup",
        "Run dotnet test to confirm all tests pass"
      ]
    }
  ]
}
```

Generated markdown (`onboarding-engineering.md`):

```markdown
# Engineering Onboarding

## Welcome to the Team

Welcome to the engineering team. This guide walks you through
your first week setup and key processes.

## Development Environment Setup

- Install .NET 10 SDK
- Clone the repository: `git clone https://github.com/org/repo`
- Run `dotnet build` to verify setup
- Run `dotnet test` to confirm all tests pass

## Code Review Process

...
```

### Try It Yourself

1. Configure pptx-mcp in your AI assistant.
2. Use a training or documentation deck you already have.
3. Ask the agent to export it to markdown using the prompt above.
4. Review the output and check it into your docs repo.

> **Note:** A dedicated `pptx_export_markdown` tool is planned for Phase 1 that will automate the assembly step, producing a single markdown string in one call.

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

## 4. Data Dashboard Updater *(Coming in Phase 2)*

### Scenario

Your team has a weekly board presentation with a metrics slide. Instead of manually updating KPI values each Monday morning, an AI agent fetches the latest numbers from your data source and updates the relevant slides automatically.

### Agent Prompt *(Planned)*

```
Fetch today's KPIs from our dashboard MCP server.
Then update the metrics slide (slide 3) in /presentations/weekly-board-update.pptx
with the new values: ARR, MRR, NRR, and new logo count.
```

### Tool Workflow *(Planned)*

1. **External MCP call** — Agent fetches live data from a dashboard or database MCP server.
2. **`pptx_list_slides`** — Identify which slide contains the metrics table.
3. **`pptx_get_slide_content`** — Inspect current placeholder structure and shape positions.
4. **`pptx_update_slide_data`** *(Phase 2)* — Update specific data fields in the slide with fresh values.
5. **`pptx_update_text`** — Update the "Last Updated" date stamp on the slide.

### Status

> **⚠️ Coming in Phase 2.** The `pptx_update_slide_data` tool is not yet implemented. The current `pptx_update_text` tool can update individual text placeholders by index today — see the [tool reference](../README.md) for details.

Phase 2 will add:
- `pptx_update_slide_data` — structured field updates driven by a data map
- Template variable support (`{{metric_name}}` placeholders) for repeatable data injection
- Multi-source composition examples (pptx-mcp + external data MCPs)

---

## Related Resources

- [README](../README.md) — Full tool reference and configuration
- [PRD](PRD.md) — Product requirements, goals, and roadmap
