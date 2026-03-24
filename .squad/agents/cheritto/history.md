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

### Slide data update tool (2026-03-16)
- `PresentationService.UpdateSlideData(...)` uses 1-based slide numbers for the write-facing MCP tool and resolves targets by case-insensitive shape name first, with zero-based text-shape index fallback.
- Preserving PowerPoint formatting means cloning the existing `TextBody` body/list style plus paragraph and run properties, then replacing only the text paragraphs instead of rebuilding the shape from scratch.
- `pptx_get_slide_content` is the discovery step for write operations: agents should inspect shape `Name` values before calling `pptx_update_slide_data` so updates stay deterministic on real decks.

### Shape targeting recovery (2026-03-16)
- `pptx_update_slide_data` failure messages should include `index:name` listings for available text-capable shapes so an agent can recover with a follow-up call instead of re-inspecting the deck.

### Markdown Export Tool (2026-03-16)
- `pptx_export_markdown` should keep tool logic thin and delegate markdown generation to `PresentationService.ExportMarkdown(...)`.
- Phase 1 markdown export excludes speaker notes even though notes are available elsewhere in `PresentationService`.
- Exported images belong in a sibling `<markdown-file>_images` folder and markdown should reference them with relative forward-slash paths for portability.
- Realistic PPTX fixtures need explicit paragraph/table/image construction in `TestPptxHelper` to validate nested bullets, tables, and embedded media.

### Phase 2 Completion (2026-03-16)
- **Issue #19 (Implement pptx_update_slide_data):** Completed and merged (PR #29)
- **Files:** 19 modified, +1975 lines
- **Core deliverable:** Dual-path shape targeting (shapeName + placeholderIndex fallback)
- **Key implementation:** `ReplaceShapeTextPreservingFormatting` method (PresentationService.cs) clones TextBody properties and paragraph/run formatting
- **Testing:** Unit tests + E2E scenario (4-slide KPI deck), format verification, PowerPoint round-trip
- **Code review:** Nate approved for production ("Ship it — production-ready")
- **Findings:** MCP SDK patterns match dotnet-mcp exactly; OpenXML approach is cleaner than MarpToPptx's explicit assignment
- **Recommendations:** Low-priority polish (size checks, validation helpers, documentation updates)
- **Result:** Phase 2 issues #15–#19 all closed, 66/66 tests passing (up from 52 end of Phase 1)

### Batch deck refresh tool (2026-03-16)
- `src/PptxMcp/Tools/PptxTools.cs` now exposes `pptx_batch_update(filePath, mutations)` as a thin MCP wrapper that returns structured JSON and keeps empty batches as a zero-count success case.
- `src/PptxMcp/Services/PresentationService.cs` batches named text mutations through one `PresentationDocument` open/save cycle, reuses the `UpdateSlideData` targeting/formatting path, and saves each touched slide once after processing the whole batch.
- Batch request/result contracts live in `src/PptxMcp/Models/BatchUpdateMutation.cs`, `BatchUpdateMutationResult.cs`, and `BatchUpdateResult.cs`.
- Compatibility validation for batch refresh now lives in `tests/PptxMcp.Tests/Services/BatchUpdateTests.cs`, which compares post-update `OpenXmlValidator` output against the baseline deck in addition to opening the file successfully.

### Round 1: Issue #34 (pptx_batch_update) Completion (2026-03-16T22:36Z)
- Implemented full `pptx_batch_update` tool per Phase 3 planning (#34)
- Strategy: single open/save cycle, reuses `UpdateSlideData` path, per-mutation result tracking
- Key design: successful mutations preserved even if later mutations fail (no rollback)
- Tests: +78 test cases (170/170 passing); compatibility validation via OpenXmlValidator + PowerPoint round-trip
- PR #44 merged after Nate code review (production-ready verdict)
- Code: +609 lines across 11 files
- Impact: Batch operations now unblock multi-slide workflows without repeated disk I/O; Phase 3 #34 complete
### Template-aware slide tools (2026-03-17)
- `pptx_add_slide_from_layout` and `pptx_duplicate_slide` use semantic placeholder keys in `Type` or `Type:Index` form (for example `Title`, `Body:1`, `Picture:2`) so agents can target template placeholders without relying on shape names.
- The new template-slide service logic keeps MCP tools thin, validates placeholder requests before mutation, clones slide-related parts recursively, and preserves layout/master inheritance by attaching the duplicated or generated slide to the correct `SlideLayoutPart`.
- `tests/PptxMcp.Tests\TemplateDeckHelper.cs` is the dedicated fixture builder for template-authoring scenarios; it creates multiple named layouts, indexed placeholders, shared image usage, and round-trip compatibility coverage for layout-based authoring.

### Table insert and update tools (Issue #36)
- `pptx_insert_table` creates DrawingML tables via GraphicFrame > Graphic > GraphicData > A.Table. The service method `InsertTable()` in `PresentationService.cs` handles row normalization (padding short rows), TableGrid column matching, unique shape IDs via `GetMaxShapeId()`, and the exact GraphicData URI (`http://schemas.openxmlformats.org/drawingml/2006/table`).
- `pptx_update_table` locates tables by name (case-insensitive, using `NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name`) or by zero-based index among `GraphicFrame` elements that contain `A.Table`. Cell text replacement preserves `TableCellProperties` (borders, fills) while rebuilding `TextBody` with proper BodyProperties + ListStyle + Paragraph structure.
- New DTOs: `TableInsertResult`, `TableUpdateResult`, `TableCellUpdate` in `src/PptxMcp/Models/`.
- Private helper `BuildTableRow()` creates rows with full cell structure; reusable for both header and data rows.
- Tool methods are thin wrappers returning JSON (matching `pptx_update_slide_data` pattern) with structured error results for file-not-found and exceptions.
- PR #46 created on branch `squad/36-table-tools`.

### Issue #75 — OpenXML upgrade blocked (2026-03-17)
- Issue requested upgrade of DocumentFormat.OpenXml from 3.4.1 to 3.5.0
- **Blocked:** Version 3.5.0 does not exist on NuGet; 3.4.1 is the latest published release as of 2026-03-17
- Verified via `dotnet package search` and NuGet gallery — no pre-release versions available either
- Commented on issue #75 with findings; no branch/PR created since there's nothing to ship
- Recommendation: revisit when 3.5.0 is actually published

### Tool consolidation — Issue #69 (2026-03-18)
- Consolidated `pptx_add_slide`, `pptx_add_slide_from_layout`, `pptx_duplicate_slide` into `pptx_manage_slides` with `ManageSlidesAction` enum (Add, AddFromLayout, Duplicate)
- Expanded `pptx_reorder_slides` to absorb `pptx_move_slide` via `ReorderSlidesAction` enum (Move, Reorder)
- All consolidated tool methods are `partial` and use `[McpMeta]` for machine-readable action lists
- `AddSlide` now returns structured JSON (`AddSlideResult`) instead of plain text; zero-based index from service converted to 1-based in tool layer
- `pptx_move_slide` and old `pptx_reorder_slides` now return structured `SlideOrderResult` JSON instead of plain text
- Test assertions updated: "Error:" prefix checks became "File not found" or structured JSON checks for tools that moved to `ExecuteToolStructured`
- `placeholderOverrides` parameter renamed to `placeholderValues` in the consolidated tool for consistency
- 6 new files, 1 deleted file (PptxTools.TemplateSlides.cs), 4 modified source files, 3 modified test files
- 377/377 tests passing, build warnings dropped from ~76 to ~42 (dead code removed)

### Phase 4 Scoping Complete (2026-03-24)
- **Status:** Tier 1 analysis tools unblocked and ready for implementation
- **Leads:** McCauley (scoping), Nate (OpenXML research)
- **Scope:** 7 GitHub issues (#80–#86) across 3 tiers; 32–38 hour estimate (2–3 weeks part-time)
- **Tier 1 (Read-Only Analysis):** #80 (file size breakdown), #81 (media analysis), #82 (unused layouts) — independent, low-risk, foundation for Tier 2
- **Tier 2 (Write Operations):** #83 (remove layouts), #84 (deduplicate media), #85 (compress images) — depend on Tier 1, require validation
- **Tier 3 (Deferred):** #86 (video analysis) — post-Phase-4 spike
- **Key ADRs:** Read-only analysis first; SkiaSharp for image compression; OpenXML validation + PowerPoint round-trip for Tier 2; SHA256 media dedup
- **Team:** Cheritto (implementation lead), Shiherlis (E2E validation), Nate (code review available)
- **OpenXML Research:** All issues highly feasible; reference patterns established in SKILL.md
- **Next Action:** Cheritto to begin Tier 1 tools (can implement in any order)
