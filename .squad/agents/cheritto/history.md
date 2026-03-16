# Project Context

- **Owner:** Jon Galloway
- **Project:** pptx-mcp — .NET 10 MCP server for PowerPoint manipulation via OpenXML SDK
- **Stack:** .NET 10, C#, ModelContextProtocol v1.1.0, DocumentFormat.OpenXml v3.3.0, xUnit v3 (MTP), Microsoft.Extensions.Hosting v10.0.5
- **Architecture:** Console app with stdio transport. Models → Services (PresentationService) → Tools (PptxTools) → MCP server
- **Key files:** src/PptxMcp/Tools/PptxTools.cs (169 lines, 7 tools), src/PptxMcp/Services/PresentationService.cs (464 lines, all OpenXML ops)
- **Build:** `dotnet build PptxMcp.slnx --configuration Release`
- **Test:** `dotnet test --solution PptxMcp.slnx --configuration Release --no-build`
- **Reference repos:** jongalloway/dotnet-mcp (MCP patterns), jongalloway/MarpToPptx (OpenXML patterns)
- **Created:** 2026-03-16

## Learnings

### Markdown export tool (2026-03-17)
- `src/PptxMcp/Tools/PptxTools.cs` keeps read-only MCP tools thin: validate file existence, call `PresentationService`, and return raw markdown or JSON strings.
- `src/PptxMcp/Services/PresentationService.cs` now owns markdown export formatting, including `## Slide N: Title` boundaries, subtitle-to-`###` mapping, nested bullet indentation, markdown table rendering, and image extraction with relative paths.
- `tests/PptxMcp.Tests/TestPptxHelper.cs` is the shared fixture builder for realistic PPTX content; it can now generate title/body text, nested bullets, tables, and embedded images for service and tool tests.
- Markdown export for Phase 1 intentionally excludes speaker notes and writes images to a sibling `<markdown-base>_images` folder so the saved `.md` file stays portable.

### Phase 2 Assignments (2026-03-16)
- **Issue #17 (cheritto assigned):** Test pptx_update_slide_data with real metric slides — validates PowerPoint compatibility and edge cases
- **Issue #15 (cheritto assigned):** E2E test multi-source update scenario — validates full composition workflow (Goal 2B)
- Dependency: Both #17 and #15 depend on #19 (core tool implementation) being complete
- Timeline: Phase 2 estimated 3–4 weeks after Phase 1 stabilization

### Talking points extraction tool (2026-03-17)
- `src/PptxMcp/Tools/PptxTools.cs` now exposes `pptx_extract_talking_points(filePath, topN = 5)` as a read-only MCP tool that returns per-slide JSON with `SlideIndex`, `Title`, and ranked `Points`.
- `src/PptxMcp/Services/PresentationService.cs` reuses slide-content extraction and ranks text candidates by placeholder type, bullet-like structure, and text quality while filtering noise markers like `Presenter Notes`, placeholder prompts, and formatting-only text.
- Title text is used as a fallback talking point for title-only slides, but slides that are otherwise just visual content return no extracted points.
- `tests/PptxMcp.Tests/TestPptxHelper.cs` is the canonical fixture builder for realistic PPTX tests; it supports title/body placeholders and embedded images for service-level integration coverage.

<!-- Append new learnings below. Each entry is something lasting about the project. -->
