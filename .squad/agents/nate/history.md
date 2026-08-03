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

### 2026-03-24: CLI Interface Research — Issue #94

**Research Scope:** Investigated dual-mode architecture, command surface design, and distribution strategy for adding CLI interface to pptx-mcp while maintaining MCP server capability.

**Key Findings:**

1. **Dual-Mode Architecture (APPROVED):**
   - Single entry point with args-based mode detection is proven pattern (Excel MCP production precedent)
   - Mode detection: `DetermineMode(args)` → route to `RunMcpServerAsync()` or `RunCliAsync()`
   - Fully backward compatible — existing MCP configs unchanged (default = MCP when no args)
   - Mode detection overhead: <1ms, zero startup penalty
   - No changes to MCP tools or protocols needed

2. **Program.cs Structure:**
   - 40-line entry point supports both modes with shared DI container
   - Critical pattern: Log everything to stderr in MCP mode (prevents stdout pollution)
   - CLI mode uses System.CommandLine for argument parsing (v2.0.0+)
   - Shared PresentationService (core logic) used by both paths

3. **21 Tools → CLI Commands Mapping:**
   - Organize by domain + verb-noun pattern: `pptx <domain> <verb>`
   - 7 command groups identified: inspect, analyze, optimize, edit, media, slides, export
   - All 21 MCP tools cleanly map to CLI commands with no conflicts
   - High-value compound commands identified: `pptx optimize` (all-in-one), `pptx report` (analysis summary)

4. **Distribution (Phased Approach):**
   - **Phase 1 (NuGet Global Tool, Week 1):** Framework-dependent (5-10MB), works cross-platform, Magick.NET native binaries auto-extracted
   - **Phase 2 (Scoop/Homebrew, Month 2):** Platform-specific manifests for non-.NET users
   - **Phase 3 (Docker, Optional):** CI/CD and containerized batch processing
   - Single-file publish NOT recommended — framework-dependent is superior (smaller, faster, native deps work better)

5. **Magick.NET Native Binaries (Zero Special Handling):**
   - Q8-AnyCPU NuGet package includes native binaries for all platforms (win-x64, linux-x64, osx-x64, osx-arm64)
   - Global tool install auto-extracts binaries to `~/.dotnet/tools/.store/pptx-mcp/<VERSION>/runtimes/<PLATFORM>/native/`
   - Runtime automatically selects correct binary — no configuration needed
   - Precedent: MiniCover, Entity Framework Core (both production tools with platform binaries)

6. **Caveats & Gotchas:**
   - Stdin/stdout collision in dual-mode → Mitigation: Log to stderr only in MCP mode
   - Argument parsing ambiguity (file named "serve") → Require explicit flags
   - Configuration file watching can cause tight polling → Disable FileSystemWatcher
   - Process exit codes: CLI expects 0/1, MCP different → Return exit codes from CLI
   - DI container lifecycle: Use `using var host` in CLI mode

7. **Impact on MCP Server (ZERO):**
   - No breaking changes — 100% backward compatible
   - MCP default behavior unchanged (no args = serve)
   - Existing integration tests pass unchanged
   - No security or performance implications

**Deliverable:** Comprehensive research report saved to `.squad/decisions/inbox/nate-cli-research.md` with:
- Recommended dual-mode architecture with Program.cs scaffold
- Complete 21-tool CLI mapping table
- Phased distribution strategy (NuGet → Scoop/Homebrew → Docker)
- .csproj configuration recommendations
- Effort estimates and GO/DEFER verdict

**Verdict:** **APPROVED FOR IMPLEMENTATION**
- Low risk (mode detection is 20 lines; no MCP changes)
- High value (unlocks scripting, batch workflows, one-off optimization)
- Battle-tested pattern (Excel MCP precedent)
- Minimal overhead (no startup cost)
- MVP effort: ~16-20 hours (1 week)
- Full surface: ~20-30 additional hours (weeks 2-3)

**Impact:** Research unblocks CLI implementation for Issue #94. Jon can approve dual-mode design and begin MVP development immediately.

### Phase 4 OpenXML Optimization Research (2026-03-24)
- **7-Issue Feasibility Analysis:** Reviewed all proposed Wave 1–3 tools for OpenXML implementation challenges, complexity, and risk
- **Code Sketches Delivered:** Provided implementation outlines for pptx_analyze_file_size, pptx_analyze_media, pptx_find_unused_layouts + future tools
- **Skill File Capture:** Documented key OpenXML patterns (OPC structure, media/relationship traversal, layout hierarchy) for team reference
- **Impact:** Research unblocked Cheritto on all three Wave 1 tools; high confidence in feasibility; no showstoppers identified
- **Pattern:** Early research phase on OpenXML patterns prevents mid-implementation surprises; skill file becomes living reference

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

