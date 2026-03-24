# Project Context

- **Owner:** Jon Galloway
- **Project:** pptx-mcp — .NET 10 MCP server for PowerPoint manipulation via OpenXML SDK
- **Stack:** .NET 10, C#, ModelContextProtocol v1.1.0, DocumentFormat.OpenXml v3.3.0, xUnit v3 (MTP), Microsoft.Extensions.Hosting v10.0.5
- **Architecture:** Console app with stdio transport. Models → Services (PresentationService) → Tools (PptxTools) → MCP server
- **7 MCP tools:** pptx_list_slides, pptx_list_layouts, pptx_add_slide, pptx_update_text, pptx_insert_image, pptx_get_slide_xml, pptx_get_slide_content
- **Build:** `dotnet build PptxMcp.slnx --configuration Release`
- **Test:** `dotnet test --solution PptxMcp.slnx --configuration Release --no-build` (MTP runner, uses `--filter-method` not `--filter`)
- **Reference repos:** jongalloway/dotnet-mcp (MCP patterns), jongalloway/MarpToPptx (OpenXML patterns)
- **Created:** 2026-03-16

## Learnings

### Phase 4 Wave Decomposition & Sequencing (2026-03-24)
- **3-Tier Wave Strategy:** Wave 1 (Core analysis: #80, #81, #82), Wave 2 (Enhancement & consolidation), Wave 3 (Polish & documentation)
- **Squad Reassignments:** Cheritto → 3-tool implementation sprint (wave 1), Shiherlis → 92-test parallel validation, Nate → OpenXML research/skill capture, Coordinator → PR review & docs
- **Risk Assessment:** Identified PAT permissions bottleneck for automated PR creation/label updates; contingency: manual intervention workflow
- **Execution:** Wave 1 complete with all deliverables on branch/PR (0 blockers on implementation, all items ready for integration)
- **Pattern Success:** Wave-based sequencing with parallel implementation/testing streams scales well for tooling work; allows PAT issues to surface early without blocking squad

### PRD Structure & Scope (2026-03-15)
- Created PRD at `docs/PRD.md` based on PR #1 bootstrap and Jon's vision
- Phase 1 (Content Reading) focuses on two high-value tools: extract talking points + export markdown
- Phase 2 (Intelligent Updates) deferred pending Phase 1 validation; planned for multi-source composition (pptx-mcp + external data MCPs)
- **Key decision:** Non-goals explicitly exclude GUI, legacy formats, and advanced design features to keep scope bounded
- **Recommended 4 GitHub issues** for Phase 1: two tool implementations, one E2E test, one docs pass
- Timeline estimate: 2–3 weeks Phase 1, 3–4 weeks Phase 2 (estimate includes +20% buffer)

### Phase 1 Issue Creation & Triage (2026-03-16)
- Created two GitHub milestones: Phase 1 (milestone #1) and Phase 2 (milestone #2)
- Created 9 labels for issue organization: `tool`, `testing`, `documentation`, `phase-1`, `phase-2`, `squad`, and squad role labels
- Created 4 GitHub issues for Phase 1 work:
  - **#6 (Cheritto):** pptx_extract_talking_points implementation
  - **#7 (Cheritto):** pptx_export_markdown implementation
  - **#8 (Shiherlis):** E2E testing with real presentations
  - **#9 (@copilot):** Phase 1 docs + example workflows
- Assignment logic: Tool work → Cheritto (backend dev charter), testing → Shiherlis (tester charter), docs → @copilot (small features with specs, auto-assigned, requires review)
- Dependency chain: #6 & #7 independent, #8 depends on both, #9 depends on both
- Decision document written to `.squad/decisions/inbox/mccauley-prd-phase1-issues.md`
- All issues reference PRD success criteria and use acceptance checklists for clarity

### Phase 1 Documentation Strategy (2026-03-17)
- Jon requested README rewrite and user-facing docs creation
- Created 5 new documentation issues, all Phase 1 milestone, assigned to @copilot:
  - **#10:** README rewrite — problem-centric structure (WHY + HOW + capabilities + use cases)
  - **#11:** TOOL_REFERENCE.md — comprehensive tool reference with JSON examples
  - **#12:** QUICKSTART.md — zero-to-working guide for new users
  - **#13:** CLIENT_SETUP.md — MCP client configuration (Claude Desktop, VS Code, CLI, local LLMs)
  - **#14:** EXAMPLES.md — real use case walkthroughs with agent prompts and tool workflows
- **Scope clarity:** Documentation is user-facing only (no internal architecture or decision docs)
- **Key decision:** Installation section notes future NuGet publishing; docs are future-proof
- **No hard dependencies** between docs; can be done in parallel. Will grow richer as Phase 1 tools complete
- Decision document written to `.squad/decisions/inbox/mccauley-phase1-documentation-issues.md`
### Phase 2 Decomposition (2026-03-16)
- Decomposed Phase 2 ("Content Writing & Intelligent Updates") into 5 GitHub issues (#15–#19)
- **Core tool work:** Issue #19 (cheritto) implements `pptx_update_slide_data` tool for data-driven slide updates (Goal 2A)
- **Testing:** Issue #17 (shiherlis) validates pptx_update_slide_data with real metric slides; Issue #15 (shiherlis) validates multi-source E2E scenario (Goal 2B)
- **Docs & examples:** Issue #18 (copilot) designs composition pattern demo; Issue #16 (copilot) updates all documentation
- **Dependency chain:** #19 → #17 → #15 + #18 → #16 (docs closes last)
- All issues routed to appropriate team members; all assigned to "Phase 2 — Content Writing & Updates" milestone
- Non-duplication verified: Phase 1 issues #6–#14 are independent; no overlap with Phase 2

### Phase 3 Planning & Feature Ranking (2026-03-17)

- Led Phase 3 planning session with Nate's input on prior-art research
- Ranked 7 features by product impact, complexity, and sequencing: batch refresh → template-aware authoring → tables → picture placeholders → notes → chart refresh → slide organization
- Consulted with Nate (per Jon directive) on prior-art research from MarpToPptx and dotnet-mcp; aligned on feasibility and implementation patterns
- Created 7 GitHub issues (#34–#40) under Phase 3 milestone with comprehensive ownership, acceptance criteria, and dependencies
- Recorded decision to continue McCauley+Nate partnership for major architectural decisions (worked well for aligned thinking, caught gotchas)
- Scope cuts: no full theme/master editing, no net-new chart authoring (refresh existing only).

### Quality & Housekeeping Phase (2026-03-18)

**Status:** Proposed (blocking gate before Phase 3 ramps up)

- Analyzed post-Phase-2 codebase: 260 tests passing, 68 build warnings, ~3,100 core LOC
- Identified three high-priority quality issues:
  1. **Code duplication:** 27× identical file-check + try-catch pattern in PptxTools.cs; 22× inline JsonSerializerOptions; 2× batch failure construction
  2. **Documentation staleness:** README tool list incomplete (table tools from Phase 2 not advertised); TOOL_REFERENCE.md and EXAMPLES.md need sync
  3. **Test coverage gaps:** Missing null/empty input validation, boundary conditions (slide indices, shape IDs, table columns), FileNotFound parameterization
- Assessed OpenXML patterns as consistent and well-established (no action needed); dead code audit found none
- **Created 8 GitHub issues (#48–#56)** under new "Quality & Housekeeping" milestone (tier 1: Q1–Q3 are blocking gates; tier 2: Q4–Q7 are polish; tier 3 deferred)
- Proposed schedule: Week 1 (Q1–Q3 unblock Phase 3), Week 2 (Q4–Q7 polish)
- Team assignments: Cheritto (Q2, Q4 refactoring), Shiherlis (Q3, Q5, Q8 testing), @copilot (Q1, Q6, Q7 docs)
- Decision document written to `.squad/decisions/inbox/mccauley-quality-phase.md`
- **Key finding:** Phase 2 delivered production-ready table operations; quality pass ensures Phase 3 starts with a clean foundation.

### 2026-03-17T06:07Z: Quality & Housekeeping Phase Finalized

- Completed full codebase quality analysis: baseline, duplication patterns, documentation gaps, test coverage assessment
- Outcome: 8 GitHub issues (#48–#56) under "Quality & Housekeeping" milestone
  - Tier 1 (Must Do): Q1–Q3 blocking gates (documentation, boilerplate extraction, validation tests)
  - Tier 2 (Should Do): Q4–Q7 polish (consolidation, boundary tests, doc sync)
  - Tier 3 (Nice to Have): Q8–Q10 deferred (parameterized tests, stress tests, squad archive)
- Decision: Commit to Tier 1 + Tier 2 (9 hours, 1–2 weeks part-time) before Phase 3 ramps up
- Team assignments aligned: Cheritto (refactoring), Shiherlis (testing), @copilot (docs), McCauley (oversight)
- Success criteria: All Tier 1+2 closed, 260+ tests, 68 or fewer warnings, docs reflect all tools, zero Phase 2 regression
- Orchestration log written to `.squad/orchestration-log/2026-03-17T0607Z-mccauley.md`
- Session log written to `.squad/log/2026-03-17T0607Z-quality-phase.md`
- Decisions merged to decisions.md; inbox files deleted

### Tool Consolidation API Design (2026-03-18)

- Designed API for issue #69 targeted tool consolidation (scoped to two surgical merges per Jon's direction)
- **Consolidation 1 — `pptx_manage_slides`:** Absorbs `pptx_add_slide`, `pptx_add_slide_from_layout`, `pptx_duplicate_slide` behind a `ManageSlidesAction` enum (Add, AddFromLayout, Duplicate)
- **Consolidation 2 — `pptx_reorder_slides` (expanded):** Absorbs `pptx_move_slide` behind a `ReorderSlidesAction` enum (Move, Reorder)
- **Pattern follows dotnet-mcp:** C# enums for action params, all action-specific params nullable with per-action validation, `[McpMeta]` for machine-readable action lists
- **Key decisions:** Required action param (no default), clean break (no stub redirects), per-action result types (not a unified union), structured JSON for all actions
- **New artifacts:** 4 model files (2 enums, 2 result records), 2 tool partial class files; old tool methods deleted
- **Risk flags for Cheritto:** Enum deserialization casing, zero-based index conversion for AddSlide, `partial` keyword requirement, test assertion updates (Add action now returns JSON not plain text)
- Decision document written to `.squad/decisions/inbox/mccauley-tool-consolidation-api.md`

### Phase 4: Presentation Optimization Planning (2026-03-19)

**Lead:** McCauley  
**Status:** Complete — Phase 4 scope finalized, GitHub milestone created, 7 issues filed

**Summary:** Defined Phase 4 ("Presentation Optimization") as a natural follow-on to Phase 3 (Deck Authoring). Jon regularly needs to shrink PowerPoint files by removing unused masters, deduplicating media, and compressing images. Phase 4 delivers this as a tier-structured suite: Tier 1 (read-only analysis, low risk), Tier 2 (write operations with OpenXML validation), Tier 3 (deferred video optimization).

**Tier 1 — Read-Only Analysis (3 issues, independent, implement first):**
- #80 (P4-1): Analyze file size breakdown — scan PPTX ZIP structure, report by category (slides, images, video/audio, masters, layouts, other)
- #81 (P4-2): List media assets — enumerate images/video/audio, detect duplicates by SHA256 hash
- #82 (P4-3): Find unused masters/layouts — cross-reference against actual usage, report space impact

**Tier 2 — Write Operations (3 issues, require Tier 1 foundation):**
- #83 (P4-4): Remove unused masters/layouts (depends on P4-3) — delete unused parts, preserve relationships, validate with OpenXmlValidator
- #84 (P4-5): Deduplicate media (depends on P4-2) — consolidate identical media to single canonical copy
- #85 (P4-6): Compress images (independent) — downscale to target DPI, format conversion (BMP/TIFF→PNG/JPEG), stats

**Tier 3 — Deferred (1 issue):**
- #86 (P4-7): Video optimization analysis only (video re-encoding deferred for future spike)

**Key Architectural Decisions:**
1. **Read-only first:** All Tier 1 analysis before any Tier 2 mutations (safe, diagnostic value, foundation for cleanup)
2. **SkiaSharp for image compression:** Chosen for cross-platform support, high quality (vs. System.Drawing.Common legacy, ImageSharp slow)
3. **OpenXML validation + PowerPoint round-trip:** All Tier 2 operations must validate before/after and pass file round-trip testing (learned from Phase 1/2: files can pass validator yet fail in PowerPoint)
4. **SHA256 content hash for dedup:** Deterministic, reliable, simple
5. **ZIP-level package scanning:** Use System.IO.Compression for size analysis (complementary to OpenXML SDK)

**GitHub Artifacts:**
- Milestone #5 "Phase 4: Presentation Optimization" created
- Labels added: `phase-4`, `analysis`, `optimization`, `media` (in addition to existing `squad`, `type:feature`, etc.)
- All 7 issues assigned to milestone #5, labeled `squad` + phase-specific labels
- Issue bodies include: technical approach, acceptance criteria, tier structure, size estimates, dependencies

**Scope Summary:**
- Estimated effort: 32 hours for Tier 1+2 (2–3 weeks part-time, one developer)
- Recommended sequence: P4-1 → P4-2 → P4-3 (Tier 1), then P4-4 (blocks on P4-3), P4-5 (blocks on P4-2), P4-6 (independent)
- Success criteria: All Tier 1+2 closed, OpenXmlValidator passes, round-trip tests pass, 260+ tests, docs updated, no Phase 3 regression

**Decision document:** `.squad/decisions/inbox/mccauley-phase4-optimization.md`

### Phase 4 Scoping Exercise Complete (2026-03-23)

**Lead:** McCauley  
**Status:** Scoping complete; ready for implementation  

**Summary:** Conducted full Phase 4 research spike ("Presentation Optimization & Media Analysis") following same scoping pattern as Phase 3.

**Research Findings:**
1. **Codebase Readiness:** Excellent. 377 tests baseline (no regression risk), 80 build warnings (acceptable), existing OpenXML patterns well-established in PresentationService.
2. **Scope Alignment:** Phase 4 naturally follows Phase 3 (deck authoring); addresses core user need: shrink file size by removing unused masters, deduplicating media, compressing images.
3. **Tier Structure:** 7 issues decomposed into Tier 1 (read-only analysis, 3 issues, low risk), Tier 2 (write operations, 3 issues, moderate risk), Tier 3 (deferred, 1 issue).

**Key Architectural Decisions:**
1. **ZIP-level package scanning** (System.IO.Compression) + OpenXML categorization (complementary analysis approach)
2. **SHA256 content hashing** for dedup (deterministic, reliable)
3. **SkiaSharp for image compression** (cross-platform > System.Drawing.Common legacy; faster than ImageSharp)
4. **OpenXmlValidator + PowerPoint round-trip testing mandatory** for all Tier 2 (experience: files pass validator but fail in PowerPoint)
5. **Master/layout relationships critical** — removing a master orphans its layouts; must validate before removal

**Research Requirements Identified (for Nate):**
- Master/layout relationship semantics (removal safety, orphan prevention)
- Media reference management (how to consolidate duplicates while preserving refs)
- SkiaSharp vs. System.Drawing.Common trade-offs (cross-platform, performance, PowerPoint compat)
- PowerPoint validation gotchas (validator sufficiency, round-trip testing strategy)

**Squad Reassignments (Recommended):**
- All tools (#80–#85) → Cheritto (backend dev, not Shiherlis as currently assigned)
- Testing for Tier 1+2 → Shiherlis (separate testing issues for each tier)
- Documentation → @copilot (add 6 tools to TOOL_REFERENCE, example workflows)
- #86 (video) → Parked (go:no, future spike)

**Critical Path (Single Developer):**
- Week 1: Tier 1 (#80, #81, #82) — 10–12 hours (sequential, independent)
- Week 2: Tier 2 (#83, #84, #85) — 16–20 hours (build on Tier 1, higher risk, needs E2E PowerPoint testing)
- Total: 32–38 hours, 2–3 weeks part-time

**Risk Flags:**
- PowerPoint compatibility (OpenXmlValidator insufficient; round-trip testing non-negotiable)
- Relationship breakage (removing parts can orphan references; comprehensive validation required)
- Image quality trade-offs (aggressive compression degrades fidelity; recommend 85% JPEG quality starting point)

**Success Criteria:**
- All Tier 1+2 issues closed
- 400+ tests passing (no Phase 1–3 regression)
- All Tier 2 E2E tests include PowerPoint round-trip validation
- Documentation updated

**Key Patterns for Reuse:**
- Tier 1 tools establish ZIP enumeration + OpenXML categorization (foundation for all downstream size reporting)
- P4-2 hash-based dedup pattern used by P4-4 (media reference consolidation) and P4-5 (image compression)
- P4-3 master/layout traversal pattern used by P4-4 (removal logic)
- All Tier 2 tools share: OpenXmlValidator before/after, error handling, detailed logging

**Decision Document:** `.squad/decisions/inbox/mccauley-phase4-scoping.md`

**Files Referenced:**
- `src/PptxMcp/Services/PresentationService.cs` (partial class pattern, OpenXML traversal)
- `src/PptxMcp/PptxMcp.csproj` (dependencies: DocumentFormat.OpenXml 3.5.1, ModelContextProtocol 1.1.0; add SkiaSharp for P4-6)
- `tests/PptxMcp.Tests/` (377 tests baseline, xUnit v3 on MTP runner)
