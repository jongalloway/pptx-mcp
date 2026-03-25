---
name: "markdown-export"
description: "Patterns for exporting PPTX presentations to markdown"
domain: "pptx-export"
confidence: "high"
source: "issue-7 implementation"
---

## Context

Use this pattern when adding or updating markdown export behavior for PowerPoint decks in `pptx-tools`.

## Patterns

### Tool Boundary
- Keep `pptx_export_markdown` thin in `src/PptxTools/Tools/PptxTools.cs`.
- Validate file existence in the tool and delegate export generation to `PresentationService.ExportMarkdown(...)`.

### Export Semantics
- Start the markdown document with a `#` heading from the first slide title when available; otherwise fall back to the file name.
- Emit each slide as `## Slide N: Title`.
- Exclude title placeholder text from slide body output to avoid duplication.
- Render subtitle placeholders as `###` headings.
- Render body paragraphs as markdown list items when placeholder/body semantics or bullet metadata indicate list content.
- Preserve nesting with two-space indentation per PowerPoint paragraph level.
- Render tables as standard markdown tables.
- Export embedded images into a sibling `<markdown-file>_images` directory and reference them using relative forward-slash paths.
- Phase 1 excludes speaker notes.

### Testing
- Use `TestPptxHelper.CreatePresentation(...)` to build realistic fixtures with multiple slides, nested bullets, tables, and images.
- Cover both service-level export behavior and MCP tool behavior.
- Validate with `dotnet build PptxTools.slnx --configuration Release` and `dotnet test --solution PptxTools.slnx --configuration Release --no-build`.

## Anti-Patterns
- Do not put OpenXML traversal or markdown formatting logic in the tool class.
- Do not emit absolute image paths in markdown output.
- Do not include speaker notes until the product decision changes.
