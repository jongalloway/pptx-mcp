# pptx-tools Product Requirements Document

## Vision

**pptx-tools** is a .NET-based Model Context Protocol (MCP) server that gives AI agents native access to PowerPoint manipulation and content extraction. It enables agentic workflows to read, analyze, update, and generate presentations programmatically through a clean, composable tool interface.

By exposing PowerPoint operations as MCP tools, pptx-tools bridges the gap between AI reasoning and Office document creation—enabling scenarios like intelligent content extraction, data-driven slide generation, and dynamic presentation updates pulling from live data sources.

---

## Current State

**PR #1 — Bootstrap pptx-mcp** established:
- **.NET 10 console application** with stdio MCP transport
- **7 MCP tools** for core PowerPoint operations:
  - `pptx_list_slides` — enumerate slides with metadata
  - `pptx_list_layouts` — list available slide layouts
  - `pptx_add_slide` — create new slides with specified layout
  - `pptx_update_text` — modify text in shapes
  - `pptx_insert_image` — embed images (PNG, JPG, GIF)
  - `pptx_get_slide_xml` — retrieve raw slide XML (power users)
  - `pptx_get_slide_content` — extract structured slide content
- **Architecture:** Models → Services (PresentationService) → Tools (PptxTools) → MCP server
- **Dependency stack:** OpenXML SDK v3.3.0, ModelContextProtocol v1.1.0, xUnit v3 on MTP
- **Test coverage:** 29 passing tests; reference implementations from jongalloway/dotnet-mcp and jongalloway/MarpToPptx
- **Status:** Ready for Phase 1 development

---

## Goals

### Phase 1 — Content Reading & Extraction (Current)
**Objective:** Enable AI agents to intelligently read and analyze PowerPoint presentations.

- **Goal 1A:** Extract key talking points from slides
  - Example: "Get me the top 5 bullet points from the .NET 10 announcement deck"
  - Build higher-level tools that parse slide content and summarize

- **Goal 1B:** Export presentations to structured markdown
  - Example: "Create a markdown file from this presentation's content"
  - Generate clean markdown output suitable for documentation or further AI processing

- **Acceptance:** Agents can read a presentation, extract meaningful content, and export to markdown without manual intervention

### Phase 2 — Content Writing & Intelligent Updates (Planned)
**Objective:** Enable agents to update presentations with live data and external context.

- **Goal 2A:** Data-driven slide updates
  - Example: "Update the metrics slide with today's latest dashboard numbers"
  - Agents fetch data and update specific slides dynamically

- **Goal 2B:** Multi-source intelligent updates
  - Example: "Update this deck based on the newest Microsoft Extensions for AI release using the Microsoft Learn MCP and the latest .NET blog post"
  - Compose multiple MCP servers (pptx-tools + Microsoft Learn MCP + web browsing) to research and update presentations

- **Acceptance:** Agents can orchestrate complex updates pulling from multiple sources and sync presentation content

---

## Use Cases

1. **Meeting Prep Assistant**
   - Agent reads your keynote deck, extracts key talking points, and summarizes by topic
   - Human presenter can quickly verify content alignment before the event

2. **Documentation Generator**
   - Agent exports a training presentation to markdown for inclusion in knowledge base
   - Non-technical content owners can maintain decks as source-of-truth; markdown stays in sync

3. **Data Dashboard Updater**
   - Agent fetches latest KPIs from a data source, updates specific slides
   - Weekly board presentations automatically refresh without manual editing

4. **Research Synthesis Tool**
   - Agent reads multiple research decks, extracts findings, creates a new summary deck
   - Enables rapid literature review and competitive analysis

5. **Blog-to-Deck Pipeline**
   - Agent reads latest blog posts via MCP, creates new presentation slide-by-deck
   - Keeps product updates and internal training decks synchronized with external communications

---

## Architecture

### Current (PR #1)
```
┌──────────────────┐
│   MCP Server     │
│  (PptxToolsServer) │
└────────┬─────────┘
         │
    ┌────▼────────────┐
    │  PptxTools      │  (MCP tool definitions)
    └────┬────────────┘
         │
    ┌────▼──────────────────┐
    │ PresentationService   │  (business logic)
    └────┬──────────────────┘
         │
    ┌────▼─────────────────────┐
    │ Models (Slide, Shape)    │  (domain objects)
    └──────────────────────────┘
         │
    ┌────▼─────────────────────┐
    │ OpenXML SDK              │  (PowerPoint I/O)
    └──────────────────────────┘
```

### Planned Extensions
- **Content analysis layer** (Phase 1): Tools to parse slide content, extract themes, generate summaries
- **Multi-source orchestration** (Phase 2): Compose pptx-tools with other MCPs to create intelligent workflows
- **Template library** (Future): Pre-built slide templates and formatting helpers

### Key Patterns
- **Stateless tools** — each tool call is self-contained; no session state
- **Clear separation of concerns** — models, services, and tools kept distinct
- **Error handling** — meaningful error messages for malformed presentations or invalid operations

---

## Non-Goals

- **GUI or interactive UI** — pptx-tools is agent-first, not user-facing
- **Legacy Office format support** (.ppt, .xls, .doc) — .NET OpenXML SDK only supports modern formats (.pptx)
- **Presentation rendering or viewing** — we manipulate and read structure, not render visuals
- **Advanced design control** — animations, transitions, and complex formatting are out of scope for Phase 1
- **Multi-document transactions** — each tool operates on a single presentation

---

## Success Criteria

### Phase 1 ✓ (Complete when)
- [ ] `pptx_extract_talking_points` tool implemented: agents can reliably extract top N bullet points per slide
- [ ] `pptx_export_markdown` tool implemented: agents can export full presentations to clean markdown
- [ ] Both tools tested with 3+ real-world presentations
- [ ] Documentation updated with Phase 1 tool examples and use cases
- [ ] Jon validates against "Get me top bullet points" and "Create markdown" scenarios

### Phase 2 ✓ (Complete when)
- [x] `pptx_update_slide_data` tool implemented: agents can update specific fields on slides
- [ ] At least one example MCP composition (pptx-tools + external data source) working
- [ ] Multi-source update scenario tested end-to-end
- [ ] Jon validates against "Update based on live data" scenario

---

## Recommended Issues for Phase 1

Create these GitHub issues to structure Phase 1 work:

1. **Tool: Extract Talking Points**
   - Title: `Implement pptx_extract_talking_points tool`
   - Description: Add tool to intelligently extract key bullet points from slide shapes. Should filter out noise (e.g., "Presenter Notes") and return structured list of key points per slide.
   - Complexity: Medium
   - Depends on: None

2. **Tool: Export to Markdown**
   - Title: `Implement pptx_export_markdown tool`
   - Description: Add tool to export full presentation structure (slides, shapes, text content) to clean markdown file. Include heading hierarchy, bullet structure, and image references.
   - Complexity: Medium
   - Depends on: None

3. **Test: Phase 1 E2E**
   - Title: `E2E test: read real presentation and export markdown`
   - Description: Integration test using real-world .pptx files. Verify extraction accuracy and markdown output quality.
   - Complexity: Low
   - Depends on: Issues 1, 2

4. **Docs: Phase 1 Examples**
   - Title: `Document Phase 1 tools and example workflows`
   - Description: Add examples to README showing how to use new tools. Include sample Agent prompts and expected outputs.
   - Complexity: Low
   - Depends on: Issues 1, 2

---

## Timeline Estimate

- **Phase 1:** 2–3 weeks (Extract Talking Points + Export Markdown + tests + docs)
- **Phase 2:** 3–4 weeks (Data updates + multi-source composition + examples)
- **Buffer:** +20% for integration complexity and edge cases

---

## Open Questions / Parking Lot

- Should Phase 1 include speaker notes extraction, or stick to visible slide content only?
- For Phase 2 data updates: should we support templating (e.g., {{metric_name}}) or just direct field replacement?
- Should we build a Python client wrapper for easier integration, or assume all consumers use the stdio MCP transport?
