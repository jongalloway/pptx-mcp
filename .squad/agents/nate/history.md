# Project Context

- **Owner:** Jon Galloway
- **Project:** pptx-mcp — .NET 10 MCP server for PowerPoint manipulation via OpenXML SDK
- **Stack:** .NET 10, C#, ModelContextProtocol v1.1.0, DocumentFormat.OpenXml v3.3.0, xUnit v3 (MTP), Microsoft.Extensions.Hosting v10.0.5
- **Reference repos:** jongalloway/dotnet-mcp (C# MCP patterns, testing, publishing), jongalloway/MarpToPptx (OpenXML PowerPoint manipulation)
- **My role:** Consulting Dev — bridge patterns from reference projects into this one
- **Created:** 2026-03-16

## Core Context (Historical Summaries)

### Phase 2 Code Review & Delivery (2026-03-16)
- Approved Cheritto's `pptx_update_slide_data` tool for production release
- MCP patterns exact match to dotnet-mcp conventions
- OpenXML text replacement via template cloning is cleaner than MarpToPptx
- Dual targeting (shapeName + placeholderIndex) perfect for multi-source composition workflows
- Phase 2 completion: All 5 issues (#15–#19) closed, PRs #29–#33 merged, 66/66 tests passing

### Batch Patterns Research & #34 Support (2026-03-16)
- Researched `IProgress<ProgressNotificationValue>` pattern from dotnet-mcp and batch strategies from MarpToPptx
- Key finding: Progress is orthogonal to error handling; both patterns are complementary
- Recommended hybrid for #34: Progress notifications + per-slide result objects + atomic PPTX save + context-rich exceptions
- Delivered comprehensive pattern guide (merged to decisions.md) with code templates and implementation checklist

### Phase 3 Planning Collaboration (2026-03-17)
- Collaborated with McCauley on Phase 3 planning per Jon directive
- Completed research on MarpToPptx prior art: template-aware authoring (proven), table writing (proven), notes writing (proven), chart authoring (net-new, no prior art)
- Identified highest-ROI path: template-aware authoring using placeholder identity and layout/master inheritance
- Recommended Phase 3 sequence: batch refresh → authoring → tables → picture placeholders → notes → chart refresh → slide organization
- Established pattern: Continue McCauley+Nate partnership for major decisions (worked well; caught gotchas)

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

### 2026-03-17: Batch Patterns Research — Issue #34 Support

**Research Scope:** Investigated `IProgress<ProgressNotificationValue>` pattern from dotnet-mcp and batch/error-handling strategies from MarpToPptx to inform Cheritto's #34 implementation (batch slide update tool).

**Key Findings:**
- **dotnet-mcp Progress Pattern:** `ExecuteWithProgress()` helper provides real-time progress reporting via MCP notifications. Pattern: report at start (Progress=0, Total=items), update per-item, report at completion (Progress=Total) even if operation throws. Null-safe (`IProgress<T>?` parameter is optional). **Critical insight:** Progress is orthogonal to error handling—it reports *state*, not *outcomes*.
- **MarpToPptx Batch Strategy:** Stop-on-first-error (fail-fast). One bad slide aborts entire render. No per-item result tracking. Rationale: PPTX atomicity (partial files can't be opened by PowerPoint). Compensates with context-rich exception wrapping (slide index + operation in message).
- **Recommended for #34:** Hybrid pattern combining both: (1) Real-time progress via `IProgress<ProgressNotificationValue>?` parameter, (2) Per-slide result objects with success/failure/message, (3) Atomic PPTX file write (all or nothing), (4) Exception wrapping for context. Tool can decide fail-on-first vs. collect-all-errors semantics in the finally block.
- **MCP Convention Alignment:** MCP SDK already defines `ProgressNotificationValue { Progress, Total, Message }` record. Use `[McpServerTool]` attribute, nullable IProgress parameter, structured JSON result.

**Deliverable:** Comprehensive pattern guide with concrete code templates, comparison table, and implementation checklist for Cheritto.

**Impact:** Unblocks #34 design phase; Cheritto has battle-tested patterns from two shipped reference projects ready to adopt.

**File Paths:**
- dotnet-mcp: `DotNetMcp/Tools/Cli/DotNetCliTools.Core.cs` (~line 178) — ExecuteWithProgress helper
- MarpToPptx: `src/MarpToPptx.Cli/Program.cs` — CLI error handling; `src/MarpToPptx.Pptx/Rendering/OpenXmlPptxRenderer.cs` — batch slide loop (referenced)## Learnings

### 2026-03-17T06:07Z: Tool Consolidation Research Integrated into Quality Pass

- Completed consolidation research finalization; framed as optional enhancement for quality pass
- Key deliverable: `.squad/decisions/inbox/nate-tool-consolidation.md` (21 KB comprehensive research)
- Consolidation opportunity: 18 tools → 6–8 consolidated (conservative: 18 → 12)
- Recommendation: Optional feature; recommend deferring to post-Tier-1 planning if pursued
- Pattern transplant ready: dotnet-mcp enum + switch routing + attribute introspection + validator pattern
- Risk mitigation: Conservative approach (slide management → text content → tables first)
- Decision point: Squad can choose to proceed with consolidation as enhancement or focus on Tier 1+2 core quality work
- Orchestration log written to `.squad/orchestration-log/2026-03-17T0607Z-nate.md`
- Decisions merged to decisions.md; inbox files deleted
