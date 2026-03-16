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
