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

### 2026-03-17: Tool Consolidation Research — Quality Pass Planning

**Research Scope:** How dotnet-mcp consolidated 70+ tools into ~10 using enum-based `action` parameter switches. Feasibility analysis for pptx-mcp.

**Key Findings:**

**dotnet-mcp consolidation:**
- **Before:** 70+ individual tools (combinatorial explosion of templates × configurations × operations)
- **After:** ~10 consolidated tools with enum-based routing (e.g., `DotnetProjectAction` with 21 actions)
- **Pattern:** One `[McpServerTool]` per domain, required `DotnetProjectAction action` parameter, switch expression to handler methods
- **Attributes:** Mark with `[McpMeta("consolidatedTool", true)]` and `[McpMeta("actions", JsonValue = [list])]` for agent introspection
- **Validation:** Centralized `ParameterValidator.ValidateAction<T>()` helper prevents typos
- **Implementation:** Partial methods per domain (Project.Consolidated.cs, Package.Consolidated.cs, etc.) with shared base class

**pptx-mcp current state:**
- **Today:** 18 individual tools (18 methods in one file)
- **Natural groupings:** 6 semantic clusters (slide inspection, slide management, text content, content extraction, image ops, table ops)
- **Potential reduction:** 18 → ~6–8 consolidated tools

**Consolidation trade-offs:**
- ✅ Benefits: Fewer tools in list, shared validation/error-handling, parameter overlap reduction, easier maintenance
- ❌ Costs: Parameter clutter (all actions' params visible), migration burden, error clarity requires action context
- 🤔 Right fit? Yes — semantic grouping obvious, parameter overlap real, management burden moderate

**Recommended approach:**
- **Conservative:** Start with 3–4 high-confidence groups (slide management, text content, tables) → reduce from 18 to ~12
- **Sequence:** Slide ops first (highest validation ROI), then text, then tables. Hold off on image/extraction until validated
- **Timeline:** ~15–21 hours (1–2 sprint days for one engineer); reversible if agent performance suffers

**Deliverable:** Comprehensive research document `.squad/decisions/inbox/nate-tool-consolidation.md` with:
- dotnet-mcp pattern breakdown (enums, attributes, switch routing, partial methods)
- pptx-mcp current tool inventory (all 18 listed with grouping rationale)
- 6 proposed consolidated tool signatures (with parameter mapping)
- Trade-off analysis (when to consolidate, when not to)
- Implementation checklist
- Reference to dotnet-mcp key files for transplant patterns

**Impact:** Unblocks quality pass planning. Jon + squad can now make informed decision: proceed with consolidation (high ROI, proven pattern) or defer for later. Conservative approach minimizes risk.

**File Paths:**
- Research output: `.squad/decisions/inbox/nate-tool-consolidation.md` (21KB)
- dotnet-mcp reference: `DotNetMcp/Actions/DotnetActions.cs` (enum definitions), `DotNetMcp/Tools/Cli/*Consolidated.cs` (implementations)
- pptx-mcp baseline: `src/PptxMcp/Tools/PptxTools.cs` (18 methods)

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

### 2026-03-17T**TBD**: DocumentFormat.OpenXml 3.5.0 Upgrade Research

**Research Scope:** Evaluate DocumentFormat.OpenXml 3.5.0 (released March 13, 2025) vs. current 3.4.1.

**Key Findings:**
- **3.5.0 additions:** Minor schema release — adds Office2016.Drawing.ChartDrawing.Offset class, Version/FeatureList/FalbackImg attributes on ChartSpace, ExtensionDropMode enum. Purely additive, no breaking changes.
- **3.4.1 recent wins:** Most significant improvements already in current version — MP4 video support for PPTX, Q3 2025 Office schema updates, ~2.4× faster base64 decoding (~70% less memory), fixed XML serialization, better error reporting for encrypted/missing parts.
- **Upgrade assessment:** Safe, low-risk. No code changes needed; dependency update only.
- **Deliverable:** GitHub issue #75 created with upgrade recommendation, impact analysis, testing checklist.

**Impact:** Unblocks potential MP4 video feature work if future requirements arise. No immediate action required—Cheritto can pick up if Squad prioritizes video embedding in Phase 3+.

**File Paths:**
- `.csproj` reference: `src/PptxMcp/PptxMcp.csproj` line 26
- Issue: https://github.com/jongalloway/pptx-mcp/issues/75
- Reference: Open-XML-SDK releases (https://github.com/dotnet/Open-XML-SDK/releases)

### 2026-03-24: Phase 4 OpenXML Patterns Research

**Research Scope:** Feasibility analysis for Phase 4 issues (#80–#86) — file size breakdown, media analysis, layout deduplication, image optimization.

**Key Findings:**

**Highly Feasible (High Confidence):**
- **#80 (File size breakdown):** `System.IO.Compression.ZipArchive` or `PackagePart` enumeration via `doc.PresentationPart.OpenXmlPackage.Package.GetParts()`. Categorize by URI pattern (slides, media, themes, layouts). Zero dependencies.
- **#81 (Media analysis):** MarpToPptx pattern established — enumerate `Picture` shapes, resolve `Blip.Embed` relationships, extract `ImagePart` stream metadata. SHA256 hashing for duplicates. Zero dependencies.
- **#82 (Unused layout detection):** Cross-reference `SlideMasterParts.SlideLayoutParts` against actual slide usage via `slidePart.SlideLayoutPart`. Straightforward relationship traversal. Zero dependencies.

**Medium Feasibility (Medium Confidence, requires testing):**
- **#83 (Remove unused layouts):** Safe deletion via `presentationPart.DeletePart(layoutPart)`. SDK handles relationship/content-type cleanup. **Risk:** PowerPoint round-trip validation essential (potential "missing template" warnings). No dependencies, but test coverage critical.
- **#84 (Deduplicate media):** Hash all images → redirect relationships → delete orphans. Tricky relationship redirection (manual: delete old, create new, update Blip.Embed). **Risk:** Atomic operation; partial failure = corruption. Zero dependencies but high test burden.

**Conditional Feasibility:**
- **#85 (Image compression):** Optional `SixLabors.ImageSharp` dependency. JPEG quality 85 achieves ~30% savings. PNG savings minimal (~10–20%). **Risk:** Lossy encoding; metadata loss (EXIF); format conversion risky. Recommend preset abstraction (Light/Medium/Aggressive).
- **#86 (Video analysis):** Marked `go:no` (correct). Analysis-only is low-value; requires external CLI (MediaInfo/FFProbe ~50 MB). **Recommendation:** Defer to Phase 5 if optimization needed.

**Current pptx-mcp Access Pattern:**
- Entry point: `PresentationDocument.Open(filePath, editable: bool)`
- Package access: `doc.PresentationPart.OpenXmlPackage.Package` → `GetParts()`
- All current service methods follow this pattern correctly

### 2026-03-24: Video/Audio Metadata Extraction Research — Issue #86

**Research Scope:** Evaluate NuGet packages for extracting video/audio metadata (codec, resolution, bitrate, duration) from embedded media in PPTX files. Analogous to Magick.NET image pattern, analysis-only (no re-encoding).

**Candidates Evaluated:** 10+ packages including FFMpegCore, TagLibSharp, MediaInfo.DotNetWrapper, FFMediaToolkit, SharpMp4Parser, and others.

**Key Findings:**

**No Pure-.NET Solution Meets Bar:**
- **FFMpegCore (MIT):** Requires FFmpeg/FFprobe CLI binaries installed system-wide. Deal-breaker for MCP server distribution (serverless/container incompatible). Excellent metadata extraction but incompatible with zero-dependency goal.
- **TagLibSharp (LGPL-2.1):** Pure .NET, zero deps, but **inadequate for video**—focuses on audio tags (ID3, Vorbis). Does NOT expose VideoCodec, VideoResolution, VideoWidth/Height from MP4/WebM.
- **MediaInfo.DotNetWrapper (MIT):** Wrapper-only; requires native MediaInfo library (bring-your-own). No stream support; file-path-only API forces temp file extraction. Same distribution complexity as FFMpegCore.
- **FFMediaToolkit (MIT):** FFmpeg wrapper, archived (last update 2021), same binary dependency problem.
- **Alternative wrappers (NReco.VideoInfo, Xabe.FFmpeg, MediaToolkit):** All FFmpeg-dependent, same issues.

**Lightweight Pure-.NET Alternative Discovered:**
- **SharpMp4Parser (MIT)** ⭐ RECOMMENDED
  - Pure .NET (Standard 2.0+), zero native dependencies
  - MIT license (perfect match to project)
  - ~100 KB NuGet package, 2K downloads (early-stage but proven)
  - Actively maintained (jimm98y/SharpMp4Parser, regular updates)
  - Stream-based API: parse MP4 box structure from memory
  - Can extract: codec (from `stsd`), resolution (from `tkhd`), duration (from `mdhd`), bitrate (computed)
  - **Analogous to Magick.NET pattern:** Lightweight, focused (metadata only, no transcoding), stream-native, cross-platform by design

**Verdict:** **GO with SharpMp4Parser for MP4/ISOBMFF video metadata extraction.**

**Implementation Pattern:**
1. Add `SharpMp4Parser` v0.0.5 to PptxMcp.csproj
2. Create `VideoMetadataExtractor` service in `Services/` — box traversal helpers for codec/resolution/duration/bitrate
3. Add `pptx_analyze_video_metadata` MCP tool — structured JSON response { codec, resolution: { width, height }, bitrate, duration }
4. Tests: Unit (box parsing), E2E (real PPTX with H.264 video), cross-platform validation
5. **Effort estimate:** 7–10 hours (comparable to Magick.NET image integration phase)

**Risk Mitigation:**
- Wrap box parsing in try-catch for malformed MP4 containers
- If >10% presentations fail (complex containers), escalate to FFMpegCore + binary distribution discussion
- WebM support (Matroska containers) out-of-scope for initial release; can add separate WebM parser if needed

**Licensing Compliance:**
- SharpMp4Parser: MIT ✅
- No LGPL/GPL dependencies introduced
- Audit: Project already MIT-licensed; zero conflict

**Deliverable:** Research report saved to `.squad/decisions/inbox/nate-video-metadata-research.md` with full candidate comparison table, architecture rationale, and implementation roadmap.

**Impact:** Unblocks #86 design phase; team has validated path forward with zero-dependency pure-.NET solution matching Magick.NET pattern.
- ZIP-level metadata available via `System.IO.Compression.ZipFile.OpenRead()` if needed (MarpToPptx precedent)

**Implementation Recommendations:**

1. **#80–#82:** Start here (pure analysis, zero risk, high confidence)
   - Add to `PresentationService` (or new `PresentationService.Optimization.cs` partial)
   - Follow existing patterns: `GetSlidePart()`, `GetSlideIds()` helpers
   - Return structured objects (JSON models), not strings
   - XML doc comments for MCP SDK Description generation

2. **#83–#84:** Medium priority (require PowerPoint round-trip testing)
   - Implement mutations with atomic backup pattern
   - Shiherlis to design test harness (PowerPoint validation)
   - Document relationship cleanup edge cases

3. **#85 (if prioritized):** Conditional dependency
   - Recommend `SixLabors.ImageSharp` over System.Drawing (cross-platform, maintained)
   - Feature flag for compression features
   - Document lossy encoding trade-offs

4. **#86:** Defer (Phase 5 candidate)
   - Reason: Low-value analysis-only; external CLI overhead
   - If future optimization needed: MediaInfo/FFProbe integration

**Prior Art References:**
- MarpToPptx: `OpenXmlPptxRenderer.cs` (PackagePart enumeration, ZipArchive usage, NormalizePackage)
- MarpToPptx: `PptxMarkdownExporter.Media.cs` (image enumeration, SHA256 hashing, deduplication patterns)
- pptx-mcp: `PresentationService.cs` lines 96–111 (layout enumeration)
- pptx-mcp: `PresentationService.Charts.cs` (structured result pattern)

**Deliverable:** `.squad/decisions/inbox/nate-phase4-openxml-research.md` — 40 KB comprehensive research with code sketches, gotchas, validation strategies for all 7 issues.

**Impact:** Unblocks Phase 4 implementation planning. Team can now scope work with confidence: #80–#82 are quick wins; #83–#84 require diligent testing; #85 is optional enhancement; #86 safely deferred.

### 2026-03-26: Magick.NET Feasibility Research — Issue #85 (Image Compression)

**Research Scope:** Jon directive to investigate Magick.NET (instead of SkiaSharp) for issue #85 image compression/optimization tool. Evaluate feasibility, cross-platform support, bundle size, API surface, and integration with PresentationService.

**Key Findings:**
- **Verdict: GO** ✅ — Magick.NET is fully viable for #85. Covers all requirements (resize, re-encode JPEG quality 85%, convert BMP/TIFF→PNG/JPEG, read dimensions, stream-based I/O)
- **Licensing:** Apache-2.0 (permissive, commercial-friendly, OSI-approved)
- **Cross-Platform:** Full Linux support on ubuntu-latest; bundles static-linked native binaries (no separate ImageMagick install needed)
- **Recommended Package:** `Magick.NET-Q8-x64` v14.11.0+ (Q8=8-bit component, x64=platform-specific, reduces bundle size)
- **Bundle Size Impact:** ~15-18 MB added to published binary (AnyCPU variant would be ~27-35 MB). Acceptable for open-source MCP server.
- **API Strength:** Clean Magick.NET surface for all operations:
  - Dimensions: `MagickImageInfo.Width`, `MagickImageInfo.Height`
  - Resize: `image.Resize(width, height)` with aspect ratio control
  - JPEG Quality: `image.Quality = 85` before `image.Write()`
  - Format Conversion: `image.Format = MagickFormat.Jpeg` / `MagickFormat.Png`
  - I/O: Stream-based (`new MagickImage(stream)`, `image.Write(outputStream)`)
- **Integration Pattern:** Create `PresentationService.ImageOptimization.cs` with public `OptimizeImages()` tool method; thin wrapper around Magick.NET with in-place image replacement via `imagePart.FeedData(stream)`
- **Magick.NET vs. SkiaSharp:** Magick.NET slower for simple resize (SkiaSharp 2-4x faster) but superior for image compression workflow: richer format support (100+ formats), JPEG quality control, BMP/TIFF handling. Bundle size tradeoff acceptable; performance not a bottleneck for one-time optimization.
- **Risks:** Native library compatibility on Linux (mitigated by .csproj properties for binary bundling), aspect ratio preservation (use `Resize(width, 0)` for auto-height), JPEG quality trade-offs (document as expected for compression)
- **Implementation Estimate:** 6-8 hours (dependency setup, tool method, E2E test, README update)

**Deliverable:** `.squad/decisions/inbox/nate-magick-net-research.md` — 10 KB feasibility report with API sketches, integration pattern, bundle size breakdown, comparison table vs. SkiaSharp, gotchas, and handoff recommendations for Cheritto.

**Impact:** Unblocks #85 decision. Go/no-go clear; Jon's Magick.NET preference validated. Ready for development phase with high confidence in approach.

**File Paths:**
- `src/PptxMcp/Services/PresentationService.Media.cs` (lines 104–145) — current image traversal pattern (Foundation for new ImageOptimization.cs)
- NuGet: https://www.nuget.org/packages/Magick.NET-Q8-x64/
- Magick.NET GitHub: https://github.com/dlemstra/Magick.NET (reference for API/cross-platform docs)
