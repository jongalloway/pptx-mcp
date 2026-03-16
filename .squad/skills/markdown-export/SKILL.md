---
name: "markdown-export"
description: "Export PPTX slide content to markdown with preserved structure"
domain: "openxml"
confidence: "high"
source: "cheritto issue #7"
---

## Context

Use this pattern when adding PowerPoint-to-markdown export logic in pptx-mcp.

## Patterns

### Slide structure
- Use the first slide title as the document `#` heading fallback, otherwise use the source filename.
- Emit each slide boundary as `## Slide N: Title`.
- Skip the title placeholder when rendering body content so slide headings are not duplicated.
- Map subtitle placeholders to `###` headings.

### Lists and paragraphs
- Read `A.Paragraph` elements directly from `Shape.TextBody`.
- Treat explicit bullet or auto-numbered paragraphs as list items.
- For body placeholders with multiple paragraphs, render them as markdown bullets even when PowerPoint relies on inherited list styling.
- Indent nested bullets with two spaces per paragraph level.

### Rich content
- Render `A.Table` content as a markdown table using the first row as the header.
- Export embedded `ImagePart` assets to a sibling `<markdown-base>_images` folder and reference them with relative paths in markdown.
- Keep Phase 1 exports limited to visible slide content; speaker notes stay out of the markdown.

## Key files
- `src/PptxMcp/Services/PresentationService.cs`
- `src/PptxMcp/Tools/PptxTools.cs`
- `tests/PptxMcp.Tests/TestPptxHelper.cs`
- `tests/PptxMcp.Tests/Services/MarkdownExportTests.cs`
