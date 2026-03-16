# Project Context

- **Owner:** Jon Galloway
- **Project:** pptx-mcp — .NET 10 MCP server for PowerPoint manipulation via OpenXML SDK
- **Stack:** .NET 10, C#, ModelContextProtocol v1.1.0, DocumentFormat.OpenXml v3.3.0, xUnit v3 (MTP), Microsoft.Extensions.Hosting v10.0.5
- **Reference repos:** jongalloway/dotnet-mcp (C# MCP patterns, testing, publishing), jongalloway/MarpToPptx (OpenXML PowerPoint manipulation)
- **My role:** Consulting Dev — bridge patterns from reference projects into this one
- **Created:** 2026-03-16

## Learnings

### 2026-03-16: Phase 2 Code Review (pptx_update_slide_data)

**Review Scope:** Phase 2 implementation — `pptx_update_slide_data` tool, `UpdateSlideData` service method, `SlideDataUpdateResult` model, MULTI_SOURCE_COMPOSITION.md, E2E and unit tests.

**Key Findings:**
- **MCP SDK Patterns:** Follows dotnet-mcp conventions exactly — `[McpServerTool]` attributes, XML doc comments for Description generation, structured JSON results, exception wrapping
- **OpenXML Text Replacement:** `ReplaceShapeTextPreservingFormatting` uses template cloning (BodyProperties, ListStyle, ParagraphProperties, RunProperties) — *cleaner* than MarpToPptx's explicit property assignment approach
- **Dual Targeting Strategy:** shapeName (primary) + placeholderIndex (fallback) with `MatchedBy` breadcrumb is excellent for multi-source composition workflows
- **Test Quality:** E2E test is realistic (4-slide KPI deck, named shapes, format verification, PowerPoint compatibility checks, Unicode). Unit tests cover edge cases.
- **Documentation:** MULTI_SOURCE_COMPOSITION.md is reference-quality — concrete examples, full JSON payloads, explains *why* the pattern works

**Recommendations (all low-to-medium priority):**
1. Update MULTI_SOURCE_COMPOSITION.md line 495–500 — remove "future" language, `pptx_update_slide_data` exists now
2. Consider package structure validation helper (relationship integrity, content types) — MarpToPptx pattern
3. Document shape name stability caveat (manual PowerPoint edits can rename shapes)
4. Add defensive size check (1000 paragraph limit) to prevent runaway agent output

**Verdict:** Production-ready. Recommendations are polish, not blockers.

**Reference Repo Patterns Applied:**
- dotnet-mcp: MCP tool registration, structured responses, error handling
- MarpToPptx: OpenXML validation patterns (OpenXmlValidator + can-be-opened checks)

**File Paths:**
- `src/PptxMcp/Tools/PptxTools.cs` (lines 94–150) — tool method
- `src/PptxMcp/Services/PresentationService.cs` (lines 185–517) — UpdateSlideData + helpers
- `src/PptxMcp/Models/SlideDataUpdateResult.cs` — result record
- `tests/PptxMcp.Tests/Tools/PptxPhase2E2eTests.cs` — E2E test
- `tests/PptxMcp.Tests/Services/PresentationServiceTests.cs` (lines 102–219) — unit tests
- `docs/MULTI_SOURCE_COMPOSITION.md` — composition guide

<!-- Append new learnings below. Each entry is something lasting about the project. -->

### Phase 2 Code Review & Completion (2026-03-16)
- **Code review of #19 implementation:** Approved for production release
- **MCP patterns:** Exact match to dotnet-mcp conventions (attributes, doc comments, error wrapping)
- **OpenXML excellence:** Template cloning approach (ReplaceShapeTextPreservingFormatting) is cleaner than MarpToPptx's explicit property assignment; preserves bullets, indentation, fonts, colors
- **Dual targeting:** shapeName (primary) + placeholderIndex (fallback) with MatchedBy breadcrumb is perfect for multi-source composition
- **Test quality:** Realistic E2E (4-slide KPI deck), comprehensive edge case coverage, format verification, PowerPoint round-trip validation
- **Phase 2 completion:** All 5 issues (#15–#19) closed, PRs #29–#33 merged, 66/66 tests passing
- **Risk assessment:** All low-risk; recommendations are polish, not blockers
- **Verdict:** Production-ready. Code quality rivals reference projects (dotnet-mcp, MarpToPptx)
