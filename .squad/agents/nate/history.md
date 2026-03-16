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

### 2026-03-17: Batch Patterns Research — Issue #34 Support

**Research Scope:** Investigated `IProgress<ProgressNotificationValue>` pattern from dotnet-mcp and batch/error-handling strategies from MarpToPptx to inform Cheritto's #34 implementation (batch slide update tool).

**Key Findings:**
- **dotnet-mcp Progress Pattern:** `ExecuteWithProgress()` helper (private method in DotNetCliTools.Core.cs) provides real-time progress reporting via MCP notifications. Pattern: report at start (Progress=0, Total=items), update per-item, report at completion (Progress=Total) even if operation throws. Null-safe (`IProgress<T>?` parameter is optional). **Critical insight:** Progress is orthogonal to error handling—it reports *state*, not *outcomes*.
- **MarpToPptx Batch Strategy:** Stop-on-first-error (fail-fast). One bad slide aborts entire render. No per-item result tracking. Rationale: PPTX atomicity (partial files can't be opened by PowerPoint). Compensates with context-rich exception wrapping (slide index + operation in message).
- **Recommended for #34:** Hybrid pattern combining both: (1) Real-time progress via `IProgress<ProgressNotificationValue>?` parameter, (2) Per-slide result objects with success/failure/message, (3) Atomic PPTX file write (all or nothing), (4) Exception wrapping for context. Tool can decide fail-on-first vs. collect-all-errors semantics in the finally block.
- **MCP Convention Alignment:** MCP SDK already defines `ProgressNotificationValue { Progress, Total, Message }` record. Use `[McpServerTool]` attribute, nullable IProgress parameter, structured JSON result.

**Deliverable:** Comprehensive pattern guide written to `.squad/decisions/inbox/nate-batch-patterns.md` with concrete code templates, comparison table, and implementation checklist for Cheritto.

**Impact:** Unblocks #34 design phase; Cheritto has battle-tested patterns from two shipped reference projects ready to adopt.

**File Paths:**
- dotnet-mcp: `DotNetMcp/Tools/Cli/DotNetCliTools.Core.cs` (~line 178) — ExecuteWithProgress helper
- MarpToPptx: `src/MarpToPptx.Cli/Program.cs` — CLI error handling; `src/MarpToPptx.Pptx/Rendering/OpenXmlPptxRenderer.cs` — batch slide loop (referenced)
- Research output: `.squad/decisions/inbox/nate-batch-patterns.md` — full pattern guide

### 2026-03-16: Phase 3 Reference Repo Research
- MarpToPptx already has strong prior art for template-aware slide creation, named layout selection, placeholder inheritance, picture-placeholder image insertion, native tables, media embedding, notes writing, transitions, backgrounds, captions, accessibility text, and remote asset resolution.
- The best MarpToPptx transplant for pptx-mcp is not basic slide/image insertion (pptx-mcp already has that now), but **template-aware authoring** built around placeholder identity (`type` + `idx`) and layout/master inheritance.
- MarpToPptx does **not** provide obvious native chart-authoring prior art; tables and diagram/SVG generation are the real reusable OpenXML patterns.
- dotnet-mcp demonstrates advanced MCP SDK patterns pptx-mcp is not using yet: prompts, resources, resource subscriptions, completion handlers, progress notifications, async task-store support, and cross-cutting telemetry filters.
- Best Phase 3 sequence from prior art: template-aware authoring first, table write/update second, notes/backgrounds/transitions third, then MCP UX improvements (resources/completions/prompts), with media embedding and Mermaid/diagram insertion after that.

### 2026-03-17: Phase 3 Planning — McCauley + Nate Collaboration

- Collaborated with McCauley on Phase 3 planning per Jon directive to consult Nate early on architectural decisions
- Completed comprehensive research on MarpToPptx and dotnet-mcp prior art:
  - **MarpToPptx:** Template-aware slide creation, placeholder resolution, layout/master inheritance, picture-placeholder insertion, native tables, media embedding, notes/transitions/backgrounds, captions/alt text, SVG diagram insertion (High feasibility; Medium complexity for most features)
  - **dotnet-mcp:** Prompts, resources, subscriptions, completions, progress notifications, async task-store, telemetry filters (High feasibility; improve agent UX but not required for core Phase 3 work)
- Identified the highest-ROI upgrade path: move from raw slide mutation to **template-aware authoring** using placeholder identity and layout/master inheritance (directly applicable to features #2–#5)
- Found strong transplant patterns for batch refresh (#1), tables (#3), notes (#5), but noted chart authoring (#6) has no direct MarpToPptx prior art (would be net-new design)
- Ranked Phase 3 sequence aligned with McCauley: batch refresh first (multiplier), authoring second (fidelity), tables third (data parity), then polish/UX, then optional media
- Recommended validation discipline from MarpToPptx (OpenXmlPackageValidator patterns) for Phase 3 test harness
- Decision: Continue McCauley+Nate partnership for major decisions; model worked well (aligned thinking, caught gotchas)

### Phase 2 Code Review & Completion (2026-03-16)
- **Code review of #19 implementation:** Approved for production release
- **MCP patterns:** Exact match to dotnet-mcp conventions (attributes, doc comments, error wrapping)
- **OpenXML excellence:** Template cloning approach (ReplaceShapeTextPreservingFormatting) is cleaner than MarpToPptx's explicit property assignment; preserves bullets, indentation, fonts, colors
- **Dual targeting:** shapeName (primary) + placeholderIndex (fallback) with MatchedBy breadcrumb is perfect for multi-source composition
- **Test quality:** Realistic E2E (4-slide KPI deck), comprehensive edge case coverage, format verification, PowerPoint round-trip validation
- **Phase 2 completion:** All 5 issues (#15–#19) closed, PRs #29–#33 merged, 66/66 tests passing
- **Risk assessment:** All low-risk; recommendations are polish, not blockers
- **Verdict:** Production-ready. Code quality rivals reference projects (dotnet-mcp, MarpToPptx)

### Round 1: Batch Patterns Research — Issue #34 Support (2026-03-16T22:36Z)
- Researched `IProgress<ProgressNotificationValue>` pattern from dotnet-mcp and batch/error-handling strategies from MarpToPptx
- Key finding: Progress is orthogonal to error handling; patterns from both repos are complementary
- dotnet-mcp: Real-time progress reporting with optional `IProgress<T>?` parameter; pattern: report at start (0), per-item, completion (even on throw)
- MarpToPptx: Stop-on-first-error with atomic file writes and context-rich exceptions (slide index + operation)
- Recommended hybrid for #34: Progress notifications + per-slide result objects + atomic PPTX save + exception wrapping
- Deliverable: Comprehensive pattern guide (merged to decisions.md) with code templates, comparison table, implementation checklist
- Code review: Approved Cheritto's PR #44 as production-ready (MCP SDK patterns match dotnet-mcp exactly)
- Impact: Cheritto had battle-tested patterns from two shipped reference projects ready to adopt; team aligned on batch semantics before implementation
