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

### Phase 1 Issue Creation (2026-03-16)
- Assigned #6 & #7 (pptx_extract_talking_points, pptx_export_markdown) for implementation
- Both are Medium complexity, can be parallelized
- Tool implementations must be integration-tested on real presentations before acceptance
- Depends on Shiherlis for E2E validation (#8) and @copilot for documentation (#9)
- All issues reference docs/PRD.md for success criteria alignment
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

### Export JSON tool (#128, 2026-07-25)
- `pptx_export_json` uses consolidated action enum pattern (Full, SlidesOnly, MetadataOnly, SchemaOnly) with `[McpMeta]` attributes
- Export models embed sub-type data in shapes: `ShapeExport` has nullable `Table`, `Image`, `Chart` properties
- `SlideExport` exposes computed `Charts`/`Images` convenience properties that aggregate from shapes (serialized in JSON, not `[JsonIgnore]`)
- `SchemaOnly` action returns schema description string without reading any file — tool handles null filePath gracefully
- Reuses `GetSlideContent`, `GetPresentationMetadata`, `GetSlideNotes`, and `ExtractChartData` from existing service partial classes
- Chart data extracted via `BuildChartLookup` keyed by shape name, then matched during shape iteration

### MCP tool descriptions (2026-05-11)
- ModelContextProtocol tool descriptions are populated from XML documentation comments only when `src/PptxTools/PptxTools.csproj` emits the XML doc file (`<GenerateDocumentationFile>true</GenerateDocumentationFile>`).
- Without the generated `src/PptxTools/bin/Release/net10.0/PptxTools.xml`, reflected MCP tools advertise empty descriptions even when the tool methods have `<summary>` comments.
- This affects consolidated tool partials such as `src/PptxTools/Tools/PptxTools.ManageMedia.cs`, `PptxTools.ManageSlides.cs`, `PptxTools.Optimization.cs`, and `PptxTools.Hyperlinks.cs`.

### MCP Tool Description Fix Implementation (2026-05-11)

**Task:** Fix MCP startup warning — `[warning] Tool pptx_manage_media does not have a description`

**Solution:** Enabled `<GenerateDocumentationFile>true</GenerateDocumentationFile>` in `src/PptxTools/PptxTools.csproj`

**Verification:** Shiherlis confirmed fix effective (fresh rebuild, clean startup, 1261/1261 tests passing)

**Decision recorded:** `.squad/decisions.md` — "Enable XML Documentation for MCP Tool Descriptions (2026-05-11)"

**Outcome:** ✅ Complete. Minimal change preserves pattern of tool descriptions in XML comments. Fixes `pptx_manage_media` and sibling tools.


