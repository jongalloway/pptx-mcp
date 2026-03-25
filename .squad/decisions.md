# Squad Decisions

## Active Decisions

### Phase 1 Issue Structure & Squad Assignments (2026-03-16)

**Lead:** McCauley  
**Status:** Active

#### Summary
Created GitHub issue structure for Phase 1 ("Content Reading & Extraction") PRD. Four issues span tool implementation, integration testing, and documentation. Assigned to squad members based on capability fit and charter alignment.

#### Issues Created
- **#6 — Implement pptx_extract_talking_points tool** (Medium complexity, Cheritto)
  - New MCP tool to extract key bullet points from slides
  - Labels: `tool`, `phase-1`, `squad:cheritto`
  - Milestone: Phase 1 — Content Reading & Extraction
  - Acceptance: Structured output, unit tests (5+ cases), integration tested
  
- **#7 — Implement pptx_export_markdown tool** (Medium complexity, Cheritto)
  - New MCP tool to export full presentation to clean markdown
  - Labels: `tool`, `phase-1`, `squad:cheritto`
  - Milestone: Phase 1 — Content Reading & Extraction
  - Acceptance: Valid markdown, unit tests (5+ cases), integration tested
  
- **#8 — E2E test: read real presentation and export markdown** (Low complexity, Shiherlis)
  - Integration test on 3+ real presentations
  - Labels: `testing`, `phase-1`, `squad:shiherlis`
  - Milestone: Phase 1 — Content Reading & Extraction
  - Dependencies: #6, #7
  - Acceptance: 3+ diverse presentations tested, accuracy verified, CI passing
  
- **#9 — Document Phase 1 tools and example workflows** (Low complexity, @copilot)
  - README updates, agent prompts, JSON examples, use case narrative
  - Labels: `documentation`, `phase-1`, `squad:copilot`
  - Milestone: Phase 1 — Content Reading & Extraction
  - Dependencies: #6, #7
  - Acceptance: 2+ example prompts, JSON request/response shown

#### Squad Assignments Rationale
- **Cheritto** → Tool implementations (#6, #7) — Backend Dev charter
- **Shiherlis** → Integration testing (#8) — Tester charter  
- **@copilot** → Documentation (#9) — Small features with specs (auto-assigned, requires review)

#### Milestones & Labels
**Milestones:**
- Phase 1 — Content Reading & Extraction (milestone #1)
- Phase 2 — Content Writing & Updates (milestone #2, created for future use)

**Labels:**
- `tool` — New MCP tool implementations
- `testing` — Test coverage and E2E testing
- `documentation` — Documentation updates
- `phase-1`, `phase-2` — Phase markers
- `squad` — Squad triage marker
- `squad:cheritto`, `squad:shiherlis`, `squad:copilot` — Team assignment labels

#### Scope Decisions
1. **No speaker notes for Phase 1** — Assume visible content only; parking lot for later
2. **Three-issue dependency chain** — Structured to avoid bottlenecks
3. **Minimum test coverage** — 5+ unit tests per tool; E2E covers 3+ presentations
4. **Real-world validation** — Both tools must be tested on real presentations before acceptance

---

### User Directive: Copilot PR Review (2026-03-16)

**By:** Jon Galloway (via Copilot)  
**Status:** Active

**Directive:** Always request a review from Copilot on each PR the squad creates. Use `--reviewer @copilot` (or equivalent) when opening PRs via `gh pr create`.

**Rationale:** User request — captured for team memory

---

### Phase 1 Documentation Issues (2026-03-17)

**Lead:** McCauley  
**Status:** Active

#### Summary
Created 5 user-facing documentation issues to support Phase 1 launch. These address Jon's request to establish a problem-centric README plus essential docs for end users integrating pptx-mcp into their MCP clients. All assigned to @copilot; no hard dependencies between issues.

#### Issues Created

All assigned to @copilot (Coding Agent), Phase 1 milestone, labels: `documentation`, `phase-1`, `squad`, `squad:copilot`:

- **#10 — Rewrite README:** Problem statement + quick install + capabilities + use cases
- **#11 — TOOL_REFERENCE.md:** Complete alphabetical tool reference with parameters, returns, and JSON examples
- **#12 — QUICKSTART.md:** Zero-to-working guide (prerequisites, install, MCP config, first command, troubleshooting)
- **#13 — CLIENT_SETUP.md:** Step-by-step client configuration (Claude Desktop, VS Code, CLI, local LLMs)
- **#14 — EXAMPLES.md:** 2–4 real use case walkthroughs with agent prompts and tool workflows

#### Squad Assignment Rationale
- **@copilot** → All 5 documentation issues
- **Rationale:** Copilot charter: small features with clear specs, autonomous work
- **Parallel execution:** No blocking dependencies; all can proceed simultaneously
- **Allows Cheritto & Shiherlis** to work on #6, #7, #8 without waiting for docs

#### Scope Decisions
1. **User-facing only** — No internal architecture, dev guides, or decision docs
2. **Shareable with external teams** — Docs should be linkable to outsiders
3. **Future-proofed** — Installation section notes future NuGet publishing
4. **Runnable examples** — All JSON examples testable, not aspirational

#### Dependencies & Ordering
No hard dependencies; soft ordering for coherence:
1. README (#10) — foundation; can deepen as Phase 1 tools land
2. QUICKSTART (#12) — parallel with README
3. CLIENT_SETUP (#13) — parallel; setup guidance
4. TOOL_REFERENCE (#11) — parallel; reference all 7 current tools (will expand)
5. EXAMPLES (#14) — parallel; grows richer with Phase 1 tools

#### Future Updates
- **When Phase 1 tools complete (#6, #7):** Add to TOOL_REFERENCE and EXAMPLES, highlight in README
- **When NuGet publishing set up:** Update QUICKSTART install section
- **When new clients tested:** Expand CLIENT_SETUP
### Copilot PR Review Directive (2026-03-16)
**By:** Jon (user directive)
**Status:** Active

All PRs created by squad members must request review from Copilot. Use `"@copilot"` (double-quoted) for reviewer assignment:
- On create: `gh pr create ... --reviewer "@copilot"`
- On existing: `gh pr edit {number} --add-reviewer "@copilot"`
Note: `--reviewer copilot` (unquoted, no @) fails with "not found".

### Ralph: Copilot Branch → PR Detection (2026-03-16)
**By:** Jon (user directive)
**Status:** Active

During Ralph's work-check cycle (Step 1 scan), Ralph must check for `copilot/*` remote branches that do **not** have a corresponding open PR. GitHub's @copilot coding agent may complete work and push to `copilot/*` branches without creating PRs.

**Detection:**
```bash
# Fetch latest remote state
git fetch --all --prune

# List copilot/* remote branches
git branch -r --list 'origin/copilot/*'

# Cross-reference against open PRs
gh pr list --state open --json headRefName --jq '.[].headRefName'
```

Any `copilot/*` branch not in the open PR list → create a PR:
- Match the branch to an issue (check branch name slug against issue titles/numbers)
- `gh pr create --base main --head {branch} --title "{issue title}" --body "Closes #{issue_number}"`
- Log the PR creation in Ralph's status report

**Priority:** Run this check in every Ralph cycle alongside the existing issue/PR scans.

### Phase 2 Decomposition (2026-03-16)
**Lead:** McCauley  
**Status:** Approved

Phase 2 ("Content Writing & Intelligent Updates") has been decomposed into **5 GitHub issues** (#15–#19), assigned to the "Phase 2 — Content Writing & Updates" milestone.

| # | Title | Owner | Type | Goal | Depends On |
|---|-------|-------|------|------|-----------|
| #19 | Implement pptx_update_slide_data tool | cheritto | Tool | 2A | Phase 1 core |
| #17 | Test pptx_update_slide_data with real metric slides | shiherlis | Testing | 2A | #19 |
| #18 | Design multi-source composition example | copilot | Docs/Example | 2B | None (parallel) |
| #15 | E2E test: multi-source update scenario | shiherlis | Testing | 2B | #19, #18 |
| #16 | Update documentation: Phase 2 tools and workflows | copilot | Docs | 2A+2B | #17, #15 |

**Dependency Chain:**
```
#19 (tool) → #17 (tool testing) ─┐
                                   ├→ #15 (E2E test)
#18 (composition example) ────────┘
                                   │
                         #16 (docs, closes last)
```

**Rationale:** Core tool (#19) unblocks testing and multi-source composition. Testing (#17, #15) validates PowerPoint compatibility and end-to-end scenarios. Documentation and examples (#18, #16) demonstrate Phase 2 capabilities and close the milestone.

---

## Phase 2 Completion Summary (2026-03-16)

**Status:** Complete — All 5 issues (#15–#19) closed, PRs #29–#33 merged.

| # | Title | Owner | PR | Status |
|---|-------|-------|----|----|
| #19 | Implement pptx_update_slide_data tool | Cheritto | #29 | ✅ Merged |
| #17 | Test pptx_update_slide_data with real metric slides | Shiherlis | #32 | ✅ Merged |
| #18 | Design multi-source composition example | Copilot | #30 | ✅ Merged |
| #15 | E2E test: multi-source update scenario | Shiherlis | #31 | ✅ Merged |
| #16 | Update documentation: Phase 2 tools and workflows | Copilot | #33 | ✅ Merged |

**Metrics:**
- Tests passing: 66/66 (up from 52 end of Phase 1)
- Tool implementation: 19 files, +1975 lines (Cheritto)
- Integration tests: 7 (Shiherlis)
- E2E tests: 1 comprehensive scenario (Shiherlis)
- Code review: Production-ready verdict (Nate)
- Documentation: Reference-quality MULTI_SOURCE_COMPOSITION.md (Copilot)

---

### Shape Targeting Strategy (2026-03-16)

**By:** Cheritto (implemented for #19)  
**Status:** Completed

**Decision:** Dual-path shape selection for `pptx_update_slide_data`:
1. Primary: `shapeName` (case-insensitive exact match)
2. Fallback: `placeholderIndex` (zero-based index across text-capable shapes)

**Rationale:**
- UX: `pptx_get_slide_content` exposes stable names; agents discover before updating
- Robustness: Fallback when names are missing or generic
- Determinism: `MatchedBy` field logs resolution path
- Recovery: Error messages list available shapes/indices for agent self-correction

**Agent Guidance:** Use `pptx_get_slide_content` first, prefer shapeName for updates, fall back to index only if needed.

---

### Phase 2 Code Review (2026-03-16)

**By:** Nate (Consulting Dev)  
**Status:** Completed, Production-Ready

**Verdict:** Ship it. Production-ready. Code quality rivals MarpToPptx patterns.

**Highlights:**
- MCP SDK patterns: Exact match to dotnet-mcp conventions
- OpenXML text replacement: Template cloning approach (cleaner than MarpToPptx)
- Dual targeting: Excellent for multi-source composition workflows
- Test quality: Realistic E2E (4-slide KPI deck, format verification, PowerPoint round-trip)
- Documentation: Reference-quality MULTI_SOURCE_COMPOSITION.md

**Recommendations (Polish, Low Priority):**
1. Remove "future" language from MULTI_SOURCE_COMPOSITION.md (tool exists now)
2. Optional: Package structure validation helper
3. Optional: Document shape name stability caveat
4. Optional: Defensive 1000-paragraph size check

**Risk Assessment:** All low (well-tested, robust fallbacks, acceptable for MCP usage).

---

### User Directive: Consult Nate Early (2026-03-16)

**By:** Jon Galloway  
**Status:** Active

Proactively bring Nate in for research and code review when his expertise (OpenXML prior art, MarpToPptx, dotnet-mcp reference repos) could inform decisions. Nate is underutilized.

**Applied in Phase 2:** Code review feedback logged above.

### Phase 3 Planning (2026-03-17)

**Lead:** McCauley & Nate  
**Status:** Complete

7 features ranked for Phase 3 ("Deck Authoring & Refresh"). Ranked by ROI, complexity, and sequencing based on McCauley's product thinking and Nate's prior-art research from MarpToPptx and dotnet-mcp.

#### McCauley's Phase 3 Call

Phase 3 should move pptx-mcp from "good at updating existing text" to "capable of authoring and refreshing whole decks inside a user's template." The highest-value work collapses multi-call agent workflows, stays inside PowerPoint compatibility lines, and makes multi-source composition feel like a first-class pattern.

#### Ranked Features

1. **Transactional batch deck refresh** (M) — Apply many mutations across multiple slides in one open/save cycle. Biggest multiplier for multi-source composition. Ownership: Cheritto (impl), Shiherlis (E2E), @copilot (docs)
2. **Template-aware slide duplication and layout population** (L) — Duplicate slides and populate placeholders by semantic identity (title/body/picture). Step from "edit current deck" to "author new slides." Ownership: Cheritto (impl), Nate (prior-art guidance), Shiherlis (round-trip), @copilot (examples)
3. **Table insert/update tools** (M) — Native table writing (cell ranges, row append/remove). Reuses MarpToPptx GraphicFrame + A.Table patterns. Ownership: Cheritto, Nate (review), Shiherlis, @copilot
4. **Picture-placeholder aware image replacement** (M) — Populate/replace images by placeholder or shape name instead of raw EMU coordinates. Ownership: Cheritto, Nate (placeholder guidance), Shiherlis (PowerPoint validation), @copilot (examples)
5. **Speaker notes and source-trace writing** (M) — Write notes with citations/URLs/assumptions. Makes multi-source composition auditable. Ownership: Cheritto, Nate (package plumbing), Shiherlis (round-trip), @copilot (workflow docs)
6. **Existing-chart data refresh** (L) — Update chart data while preserving type/styling. Scope cut: refresh only, not full authoring. Ownership: Cheritto (spike + impl), Nate (research), Shiherlis (PowerPoint validation), @copilot (docs)
7. **Slide organization operations** (M) — Move, reorder, delete to assemble clean final deck. Ownership: Cheritto, Shiherlis (E2E), @copilot (docs)

#### Scope Cuts

- Do not make full theme/master editing a Phase 3 goal. Use existing layouts and template slides instead.
- Do not promise net-new chart authoring in Phase 3. If chart work ships, keep it scoped to updating existing charts only.

#### Nate's Research (MarpToPptx & dotnet-mcp)

**From MarpToPptx:**
- Template-aware slide creation with placeholder resolution and layout/master inheritance is feasible (Medium complexity, High feasibility). Reuse SlideTemplateSelector pattern.
- Picture-placeholder image insertion is proven (Medium complexity, High feasibility). Reuse AddImageIntoPicturePlaceholder pattern.
- Native table authoring via GraphicFrame + A.Table is proven (Medium complexity, High feasibility). Reuse AddTable/CreateTableCell patterns.
- Speaker notes writing via notes-master creation is proven (Medium complexity, High feasibility). Reuse AddNotesSlide/EnsureNotesMasterPart pattern.
- Embedded video/audio is proven but more delicate (Medium-High complexity, Medium feasibility).
- Mermaid/diagram insertion via SVG + blips is proven but requires new dependencies (High complexity, Medium feasibility).
- Chart authoring is NOT proven in MarpToPptx; would be net-new design work (High complexity, Unknown feasibility).

**From dotnet-mcp:**
- Prompts, resources, completions, progress notifications, async task-store, and telemetry filters are available patterns. Not required for Phase 3 core work but improve agent UX afterward.

#### Recommended Sequencing

1. Batch refresh (multiplier for multi-source composition)
2. Template-aware authoring (unlocks slide creation with fidelity)
3. Tables (read/write parity for data-driven decks)
4. Picture placeholders (template-compatible image placement)
5. Speaker notes (auditability + composition workflow polish)
6. Chart refresh (already-designed charts only)
7. Slide organization (cleanup after generation/duplication)
Then: MCP UX improvements (resources/prompts/completions), media embedding, Mermaid insertion.

#### GitHub Issues

Created 7 issues (#34–#40) under "Phase 3: Deck Authoring" milestone with `phase-3` label. No Phase 3 work should begin until Phase 2 is fully stable (all tests passing, PRs merged).

#### Decision: Phase 3 Partner Pattern

Both McCauley and Nate contributed: McCauley on feature ranking and product vision, Nate on prior-art research and feasibility validation. This pattern worked well (aligned on sequencing, caught gotchas). Continue this for major architectural decisions.

---

### Batch Update Semantics for Issue #34 (2026-03-17)

**Lead:** Cheritto (implementation), Nate (research)  
**Status:** Complete

#### Decision
- `pptx_batch_update` reuses the existing `pptx_update_slide_data` shape-resolution and formatting-preservation path for each mutation
- Batch execution keeps successful mutations even if later mutations fail; no rollback of prior successes
- Service opens the `.pptx` once, applies all mutations in memory, then saves each touched slide part once at the end

#### Rationale
Aligns batch behavior with single-update tool while avoiding repeated open/save cycles. Gives agents deterministic per-mutation recovery details without sacrificing performance or PowerPoint compatibility.

---

### Batch Processing Patterns Research (2026-03-17)

**Lead:** Nate (Consulting Dev)  
**Status:** Complete

#### Research Scope
Investigated `IProgress<ProgressNotificationValue>` pattern from dotnet-mcp and batch/error-handling strategies from MarpToPptx to inform Cheritto's #34 implementation.

#### Key Findings

1. **dotnet-mcp Progress Pattern:**
   - `ExecuteWithProgress()` helper provides real-time progress reporting via MCP notifications
   - Pattern: report at start (Progress=0, Total=items), update per-item, report at completion (Progress=Total) even if operation throws
   - Null-safe: `IProgress<T>?` parameter is optional
   - **Critical insight:** Progress is orthogonal to error handling—it reports *state*, not *outcomes*

2. **MarpToPptx Batch Strategy:**
   - Stop-on-first-error (fail-fast)
   - One bad slide aborts entire render
   - No per-item result tracking
   - Rationale: PPTX atomicity (partial files can't be opened by PowerPoint)
   - Compensates with context-rich exception wrapping (slide index + operation in message)

3. **Recommended for #34 (Hybrid Pattern):**
   - Real-time progress via `IProgress<ProgressNotificationValue>?` parameter
   - Per-slide result objects with success/failure/message
   - Atomic PPTX file write (all or nothing)
   - Exception wrapping for context
   - Tool can decide fail-on-first vs. collect-all-errors semantics

#### MCP Convention Alignment
- MCP SDK already defines `ProgressNotificationValue { Progress, Total, Message }` record
- Use `[McpServerTool]` attribute, nullable IProgress parameter, structured JSON result
- Follows dotnet-mcp patterns exactly

#### Implementation Delivered
- Cheritto implemented per research guidance
- PR #44 merged; production-ready verdict
- 170/170 tests passing

---

### Quality & Housekeeping Phase (2026-03-17)

**Lead:** McCauley  
**Status:** Proposed (blocking gate before Phase 3 ramps up)

#### Codebase Baseline
- **Build:** ✅ Passes
- **Tests:** ✅ 260 passing, 0 failures (9s 341ms)
- **Build Warnings:** 68 (CS8602 in tests — acceptable per design choice)
- **Code Size:** ~3,100 lines core + ~3,300 lines tests

#### Quality Findings

**1. Code Duplication — MEDIUM Priority**

Three overlapping patterns:
- **Pattern 1:** 27× identical file-check + try-catch boilerplate in `PptxTools.cs` (every public method)
  - Fix: Extract to private helper `ExecuteOperation<T>(filePath, op)`
  - Impact: Save ~15–20% LOC, improve consistency
  
- **Pattern 2:** 22× inline `JsonSerializerOptions { WriteIndented = true }` recreations
  - Fix: Static readonly field `JsonOptions`
  - Impact: Memory efficiency, consistency guarantee
  
- **Pattern 3:** 2× batch failure construction (file-not-found vs. exception scenarios)
  - Fix: Extract helper `MutationsToFailures(...)`
  - Impact: DRY principle, easier updates

**2. Documentation Staleness — MEDIUM Priority**

- README tool list incomplete (7 old tools; table tools from Phase 2 not advertised)
- TOOL_REFERENCE.md needs table tool documentation
- EXAMPLES.md missing table insert/update examples
- **Impact:** User confusion; first-time UX friction

**3. Test Coverage Gaps — MEDIUM Priority**

- **Gap 1:** No null/empty input validation (null shape names, empty headers[], malformed JSON)
  - Effort: 3 test methods, ~50 LOC
  - Impact: Runtime NPEs if agent sends bad input
  
- **Gap 2:** Boundary conditions (slide index == slideCount, shape ID wraparound, table column mismatch)
  - Effort: 4 test methods, ~80 LOC
  - Impact: Silent failures or PPTX corruption on edges
  
- **Gap 3:** FileNotFound tests can be consolidated with `[Theory]` parameterization
  - Effort: 2 hours
  - Impact: Reduce 15+ nearly-identical test methods
  
- **Gap 4:** Stress/performance (100+ slides, 1000+ shapes, 50+ row tables)
  - Effort: 2 test methods, ~60 LOC
  - Impact: Unknown performance characteristics; potential memory leaks

**4. Build Warnings — LOW Priority**

- 68 warnings, all CS8602 (nullable dereference) in test files
- Root cause: Test setup dereferences without null checks
- **Recommendation:** Accept + document in `.editorconfig` (test maintainability trumps warning elimination)

**5. OpenXML Patterns — Consistent ✅**
- Shape access, ID generation, table handling, error flow all well-established
- No action needed

**6. Dead Code — None Found ✅**

#### Proposed Work Items (8 GitHub Issues: #48–#56)

**Tier 1 — Must Do (Quality Gate)**

| ID | Title | Effort | Owner | Priority |
|----|-------|--------|-------|----------|
| Q1 | README tool list sync: add table insert/update docs | 30m | @copilot | 🔴 High |
| Q2 | Extract PptxTools error-handling boilerplate | 2h | Cheritto | 🔴 High |
| Q3 | Add null/empty input validation tests | 3h | Shiherlis | 🔴 High |

**Why Tier 1:** Unblocks Phase 3; correct documentation essential; test gaps are known risks.

**Tier 2 — Should Do (Polish)**

| ID | Title | Effort | Owner | Priority |
|----|-------|--------|-------|----------|
| Q4 | Consolidate JsonSerializerOptions and batch failures | 2h | Cheritto | 🟡 Medium |
| Q5 | Add boundary condition tests | 3h | Shiherlis | 🟡 Medium |
| Q6 | Update docs/TOOL_REFERENCE.md with new tools | 1h | @copilot | 🟡 Medium |
| Q7 | Add table examples to docs/EXAMPLES.md | 1.5h | @copilot | 🟡 Medium |

**Why Tier 2:** Improves test signal; reduces cognitive load; completes docs. Not blockers but worth doing before Phase 3 ramps up.

**Tier 3 — Nice to Have (Future)**

| ID | Title | Effort | Owner | Priority |
|----|-------|--------|-------|----------|
| Q8 | Refactor FileNotFound tests to use `[Theory]` | 1h | Shiherlis | 🔵 Low |
| Q9 | Add stress tests (100+ slides, 1000+ shapes, 50+ rows) | 4h | Shiherlis | 🔵 Low |
| Q10 | Archive `.squad/` to separate branch/history | 2h | McCauley | 🔵 Low |

**Why Tier 3:** Low risk of regression; can be deferred post-Phase-3.

#### Decision: Scoping for This Phase

**Recommendation:** Commit to Tier 1 + Tier 2 (9 hours total, ~1–2 weeks part-time)

**Rationale:**
- Tier 1 is blocking gate: incomplete docs + boilerplate will slow Phase 3 onboarding
- Tier 2 is debt prevention: small tests now prevent big bugs later
- Tier 3 is polish: deferred unless timeline is loose
- **Estimated impact:** 25% fewer regressions, 15% faster code review, better new-engineer onboarding

#### Success Criteria

✅ All Tier 1 + Tier 2 issues closed before Phase 3 ramps up  
✅ Test count stable or increasing (260+ tests)  
✅ Build warnings same or fewer (68 or less)  
✅ README/docs reflect all current tools  
✅ No regression in Phase 2 functionality (all round-trip tests pass)  

#### Timeline

- **Week 1:** Q1–Q3 (unblock Phase 3 start)
- **Week 2:** Q4–Q7 (polish before full feature work)
- **Backlog:** Q8–Q10 (future or parallel to Phase 3 lighter tasks)

#### Team Assignments

- **Cheritto:** Q2, Q4 (code refactoring — Backend Dev charter)
- **Shiherlis:** Q3, Q5, Q8 (test writing — Tester charter)
- **@copilot:** Q1, Q6, Q7 (documentation — Coding Agent charter)
- **McCauley:** Milestone oversight; final review before Phase 3 starts

---

### Tool Consolidation Research (2026-03-17)

**Lead:** Nate  
**Status:** Proposed (optional enhancement for quality pass)

#### Research Scope

How dotnet-mcp consolidated 70+ tools into ~10 using enum-based action parameter switches. Feasibility analysis for pptx-mcp.

#### Key Findings

**dotnet-mcp Consolidation Pattern:**
- **Before:** 70+ individual tools (combinatorial explosion)
- **After:** ~10 consolidated tools with enum-based routing (e.g., `DotnetProjectAction` with 21 actions)
- **Pattern:** One `[McpServerTool]` per domain, required `action` parameter, switch expression to handlers
- **Attributes:** `[McpMeta("consolidatedTool", true)]` + `[McpMeta("actions", [...])]` for agent introspection
- **Validation:** Centralized `ParameterValidator.ValidateAction<T>()` prevents typos
- **Implementation:** Partial methods per domain with shared base class

**pptx-mcp Current State:**
- **Today:** 18 individual tools (18 methods in one file)
- **Natural groupings:** 6 semantic clusters (slide inspection, slide management, text content, content extraction, image ops, table ops)
- **Potential reduction:** 18 → ~6–8 consolidated tools (conservative: 18 → 12)

#### Trade-Off Analysis

**Benefits:**
- ✅ Fewer tools in agent's tool list (cleaner UX)
- ✅ Shared validation/error-handling logic
- ✅ Parameter overlap reduction
- ✅ Easier maintenance

**Costs:**
- ❌ Parameter clutter (all actions' params visible)
- ❌ Migration burden (existing workflows need updates)
- ❌ Error clarity requires action context

**Fit Assessment:** YES — semantic grouping obvious, parameter overlap real, management burden moderate.

#### Recommended Approach

**Conservative Sequence (1–2 sprint days):**
1. Start with 3–4 high-confidence groups: slide management, text content, tables
2. Achieve 18 → 12 reduction first
3. Hold image + extraction until validated
4. Reversible if agent performance suffers

#### Decision Point

**Question:** Should pptx-mcp consolidate as part of quality pass?

**Recommendation:** Frame as optional enhancement. Conservative approach minimizes risk. Can defer if squad prioritizes other quality items. If proceeding, recommend post-Tier-1 planning.

---

## Governance

- All meaningful changes require team consensus
- Document architectural decisions here
- Keep history focused on work, decisions focused on direction

---

## Tool Consolidation: Slide Creation & Ordering (Issue #69)

**Date:** 2026-03-18  
**Lead:** McCauley (design), Cheritto (implementation), McCauley (review)  
**Status:** COMPLETED (PR #76 merged)  
**Scope:** Two surgical tool merges only

### API Design Summary

Consolidated four tools into two:

1. **pptx_manage_slides** — absorbs:
   - pptx_add_slide (removed)
   - pptx_add_slide_from_layout (removed)
   - pptx_duplicate_slide (removed)

2. **pptx_reorder_slides** (expanded) — absorbs:
   - pptx_move_slide (removed)

### Key Decisions

- **Action Dispatch:** C# enum-based (required, non-nullable). SDK validates at protocol level.
- **Parameter Design:** All action-specific parameters nullable with default 
ull; validated in switch branches.
- **Backward Compatibility:** Clean break — old tool methods removed entirely. MCP clients refresh on reconnect. No redirect stubs needed (project not announced yet).
- **Service Layer:** Untouched — PresentationService methods unchanged.
- **Result Types:** Action-specific records:
  - AddSlideResult — new, for Add action
  - SlideOrderResult — new, for Move/Reorder actions
  - Existing AddSlideFromLayoutResult, DuplicateSlideResult reused as-is
- **Zero-Based Index Fix:** AddSlideResult returns 1-based slide number (
ewIndex + 1), fixing inconsistency in old pptx_add_slide.
- **Tool Metadata:** XML doc <summary> lists all actions; [McpMeta("actions")] provides machine-readable action array.

### Implementation

**New files created:**
- src/PptxMcp/Models/ManageSlidesAction.cs
- src/PptxMcp/Models/ReorderSlidesAction.cs
- src/PptxMcp/Models/AddSlideResult.cs
- src/PptxMcp/Models/SlideOrderResult.cs
- src/PptxMcp/Tools/PptxTools.ManageSlides.cs
- src/PptxMcp/Tools/PptxTools.ReorderSlides.cs

**Modified files:**
- src/PptxMcp/Tools/PptxTools.cs (removed old methods)
- src/PptxMcp/Tools/PptxTools.TemplateSlides.cs (removed old methods)
- Test suite updated
- README.md updated

**Test Results:** All 377 tests passing.

**Review Verdict:** APPROVED (all 9 checklist items passed).

**Merge Status:** PR #76 squash-merged. Issue #69 auto-closed.

---

## User Directive: Backward Compatibility Scope

**Date:** 2026-03-18 09:22Z  
**Source:** Jon Galloway (via Copilot)  
**Decision:** No backward compatibility required for tool consolidation (#69).

**Rationale:**
- Project not announced publicly yet
- No external users
- Clean break acceptable

**Implication:** Old tool method removals (stubs/redirects) unnecessary. MCP clients refresh on reconnect.

---

## DocumentFormat.OpenXml 3.5.0 Upgrade Research

**Date:** 2026-03-17  
**Author:** Nate (Consulting Dev)  
**Status:** Recommendation for Squad Review  
**Recommendation:** ✅ Safe to upgrade (3.4.1 → 3.5.0). No breaking changes. No code changes required.

### Findings

**3.5.0 Release (Mar 13, 2025):**
- Scope: Minor release, purely schema/namespace additions
- Additions: Office2016 chart attributes, ExtensionDropMode enum
- Breaking Changes: None

**3.4.1 Benefits (already in use):**
- Performance: ~2.4× faster base64 decoding, ~70% less memory allocation
- Features: MP4 video support (MediaDataPartType.Mp4)
- Schema: Updated to Q3 2025 Office release
- Reliability: Fixed XML serialization, better error messages

### Upgrade Path

1. Update src/PptxMcp/PptxMcp.csproj line 26
2. Run dotnet restore
3. Run full test suite
4. Commit

### Risk Assessment

| Category | Level |
|----------|-------|
| Breaking Changes | ✅ None |
| PPTX Compatibility | ✅ Low |
| Test Coverage | ✅ Low |
| Performance | ✅ None |

**Timeline:** 15–20 minutes

**Decision Point:** Can approve as Tier 1 polish or defer to later sprint


### Phase 4: Presentation Optimization & Media Analysis (2026-03-23 McCauley)

**Lead:** McCauley  
**Status:** Ready for Implementation

#### Summary

Phase 4 ("Presentation Optimization") is a natural follow-on to Phase 3 (Deck Authoring & Media). It addresses a core user need: reducing file size by identifying and removing bloat (unused masters/layouts, duplicate media, oversized images).

#### Phase 4 Scoping: Presentation Optimization & Media Analysis

**Lead:** McCauley  
**Date:** 2026-03-23  
**Status:** Ready for Implementation  

---

## Executive Summary

Phase 4 ("Presentation Optimization") is a natural follow-on to Phase 3 (Deck Authoring & Media). It addresses a core user need: reducing file size by identifying and removing bloat (unused masters/layouts, duplicate media, oversized images). 

**Scope:** 7 GitHub issues (#80–#86) across 3 tiers:
- **Tier 1** (Read-Only, Low Risk): 3 analysis tools — foundation for everything else
- **Tier 2** (Write Operations, Moderate Risk): 3 optimization tools — depends on Tier 1, requires PowerPoint validation
- **Tier 3** (Deferred): 1 analysis tool (video metadata only; re-encoding deferred)

**Effort Estimate:** 32–38 hours (Tier 1+2), ~2–3 weeks part-time, single developer
**Risk Level:** Moderate (write operations require strict validation; PowerPoint compatibility is the success criterion)
**Start Condition:** Tier 1 issues must clear `go:needs-research`

---

## Issue Breakdown

### Tier 1: Read-Only Analysis (Independent, Implement First)

These are **low-risk, high-value diagnostic tools**. All independently scoped; no internal dependencies. Together they provide the data foundation for Tier 2 optimizations.

#### **#80 — P4-1: Analyze presentation file size breakdown**
- **Goal:** Scan PPTX ZIP structure and report sizes by category (slides, images, video/audio, masters, layouts, other)
- **Scope:** ZIP enumeration via `System.IO.Compression`, OpenXML categorization, structured JSON output
- **Key Pattern:** Will be reused by Tier 2 tools to confirm space savings
- **Size Estimate:** 3–4 hours
- **Risk:** Low (read-only, no mutation)
- **PowerPoint Compat Risk:** None

#### **#81 — P4-2: List and analyze media assets**
- **Goal:** Enumerate all media (images, video, audio), compute SHA256 hash per part, detect duplicates, cross-ref to slides
- **Scope:** ImagePart, VideoFromFile, AudioFromFile traversal; hash-based dedup analysis
- **Key Pattern:** SHA256 content hash is reused in #84; media enumeration is reused in #85
- **Size Estimate:** 3–4 hours
- **Risk:** Low (read-only, no mutation)
- **PowerPoint Compat Risk:** None

#### **#82 — P4-3: Find unused slide masters and layouts**
- **Goal:** Enumerate SlideMasterParts and SlideLayoutParts; cross-ref against actual slide usage; report unused items
- **Scope:** Master/layout traversal via OpenXML; usage cross-ref; size calculation
- **Key Pattern:** Forms foundation for #83 (removal logic)
- **Size Estimate:** 3–4 hours
- **Risk:** Low (read-only, no mutation)
- **PowerPoint Compat Risk:** None
- **Critical Note:** Warn if removing a master would orphan layouts (relationships matter)

---

### Tier 2: Write Operations (Depend on Tier 1)

These **mutate the PPTX package** and carry higher risk. All require `OpenXmlValidator` before/after validation and **PowerPoint round-trip testing** (save, open in PowerPoint, verify no corruption).

#### **#83 — P4-4: Remove unused slide masters and layouts**
- **Goal:** Delete unused master/layout parts while preserving relationship integrity
- **Depends On:** P4-3 analysis (identifies unused items); P4-1 patterns (size confirmation)
- **Scope:** Target removal, orphan prevention, relationship cleanup, OpenXmlValidator
- **Key Implementation Detail:** Cannot remove a master if it has layouts that are in use. Must validate before attempting removal.
- **Size Estimate:** 5–6 hours
- **Risk:** **MODERATE** — removing parts can break relationships if not careful. OpenXmlValidator is necessary but insufficient (files pass validator but fail in PowerPoint).
- **PowerPoint Compat Risk:** **CRITICAL** — E2E test mandatory: remove masters, save, open in PowerPoint, verify no visual corruption or errors
- **Recommended Testing:** Test with presentations of varying complexity (minimal masters, complex themes with many layouts)

#### **#84 — P4-5: Deduplicate identical media**
- **Goal:** Find identical media by hash, consolidate to single canonical copy, update all references, remove orphans
- **Depends On:** P4-2 media analysis (hash, dedup detection)
- **Scope:** Relationship update, orphan cleanup, OpenXmlValidator
- **Key Implementation Detail:** When replacing references, must ensure all relationship IDs point to canonical. Every image/video/audio reference across all slides must be updated.
- **Size Estimate:** 5–6 hours
- **Risk:** **MODERATE** — relationship management is error-prone. One missed reference breaks slide rendering.
- **PowerPoint Compat Risk:** **CRITICAL** — E2E test mandatory: dedup a presentation with duplicated images, open in PowerPoint, verify all images display correctly and identically
- **Recommended Testing:** Create test presentations with known duplicate images in different slides; verify post-dedup that all references resolve

#### **#85 — P4-6: Compress/optimize images**
- **Goal:** Downscale images larger than needed for display, optionally convert formats (BMP/TIFF→PNG/JPEG), re-encode with compression
- **Scope:** Image part enumeration, pixel vs. display size analysis, downscaling/re-encoding, format conversion, OpenXmlValidator
- **Key Architectural Decision:** Use **SkiaSharp** (cross-platform, high quality, modern) instead of System.Drawing.Common (legacy) or ImageSharp (slower)
  - **Pro:** Works on Windows/Linux/macOS; better quality than System.Drawing.Common; faster than ImageSharp
  - **Con:** New NuGet dependency (add to .csproj)
- **Key Implementation Detail:** Display size is measured in EMU (English Metric Units); must convert to pixels to determine optimal DPI
- **Size Estimate:** 6–8 hours (includes dependency research + image processing implementation)
- **Risk:** **MODERATE** — image quality is subjective; aggressive compression may degrade visual fidelity
- **PowerPoint Compat Risk:** **CRITICAL** — E2E test mandatory: compress a presentation with mixed image types/sizes, open in PowerPoint, verify visual quality acceptable and file size reduced
- **Recommended Testing:** Create test presentations with large images, oversized photos, mixed formats; verify post-compression that images display at correct size with acceptable quality
- **Quality Tuning:** Document JPEG quality trade-offs (recommend 85% as starting point)

---

### Tier 3: Deferred (Future Spike)

#### **#86 — P4-7: Optimize embedded video (Analysis Only)**
- **Goal:** Detect video metadata (codec, resolution, bitrate, duration) and suggest compression; **no transformation**
- **Status:** Marked `go:no` — parking lot for future spike
- **Rationale:** Video re-encoding requires external tool (ffmpeg, external API) and is complex. Analysis-only is quick (2–3 hours) but low immediate value without re-encoding. Better to deliver Tier 1+2 first, then circle back if users demand video optimization.
- **Future Work:** If brought back, would likely require ffmpeg integration or external video optimization service API

---

## Implementation Sequence & Dependencies

### Critical Path (Blocking Order)

```
Tier 1 (All Parallel):
  [#80] —┐
  [#81] —├→ All Tier 1 independent
  [#82] —┘

Tier 2 (Sequential, with dependencies):
  [#82 complete] → [#83] (Remove masters depends on #82 analysis)
  [#81 complete] → [#84] (Dedup media depends on #81 analysis)
  [No dependency] → [#85] (Image compression independent, pairs well with #81 but doesn't block)
```

### Recommended Sequence for Single Developer

**Week 1 (Tier 1 Analysis — 10–12 hours):**
1. **#80** (P4-1) — File size breakdown: 3–4 hours
   - Establish ZIP enumeration + OpenXML categorization patterns
   - Used by all downstream tools for verification
2. **#81** (P4-2) — Media enumeration + dedup: 3–4 hours
   - Establishes hash-based media analysis pattern
   - Required by #84 and #85
3. **#82** (P4-3) — Master/layout finder: 3–4 hours
   - Establishes master/layout traversal patterns
   - Required by #83

**Week 2 (Tier 2 Optimization — 16–20 hours):**
1. **#83** (P4-4) — Remove masters/layouts: 5–6 hours
   - Build on #82 analysis logic
   - Highest risk; needs rigorous testing (PowerPoint round-trip)
2. **#84** (P4-5) — Dedup media: 5–6 hours
   - Build on #81 hash analysis
   - Also high risk; relationship management critical
3. **#85** (P4-6) — Image compression: 6–8 hours
   - Independent implementation; add SkiaSharp dependency
   - Can start in parallel with #84 if testing allows

**Tier 3 (Parked for future):**
- #86 deferred; not starting Phase 4

---

## Research Requirements (Before `go:needs-research` → `go:ready`)

### For Nate (OpenXML Patterns & Prior Art)

1. **Master/Layout Relationship Semantics**
   - How do OpenXML relationships between masters and layouts work? Confirm that a layout can have only one parent master, and removing the master orphans the layout.
   - Reference: MarpToPptx and dotnet-mcp precedent — did these repos handle master removal? What gotchas did they encounter?
   - **Deliverable:** Short research note on master/layout removal safety (what to validate, what can break)

2. **Media Reference Management in Slides**
   - How do image relationships work when multiple slides reference the same media? When consolidating duplicates, which relationship structure is canonical?
   - **Deliverable:** Guidance on how to safely update image references without orphaning parts

3. **SkiaSharp vs. System.Drawing.Common Trade-offs**
   - Document cross-platform support, performance, and PowerPoint compatibility for each option
   - **Deliverable:** Recommendation on which library to use for P4-6 (image compression)

4. **PowerPoint Validation & Round-Trip Testing**
   - OpenXmlValidator passes but files still fail in PowerPoint — known gotchas?
   - What's the safest way to test after removing parts or updating relationships?
   - **Deliverable:** Test strategy guidance for Tier 2 operations

### For Cheritto (Implementation Guidance)

1. **Tier 1 tools are straightforward** — focus on clean JSON output and reusable patterns (will be referenced by Tier 2)
2. **Tier 2 tools require:**
   - Comprehensive error handling (removal failures, orphan prevention)
   - Detailed logging (before/after sizes, what was removed, validation results)
   - Unit tests + E2E tests (at minimum, E2E must include PowerPoint round-trip on real files)
3. **SkiaSharp Integration** (P4-6 specific)
   - Add to `.csproj` early
   - Research DPI calculation (EMU → pixels → DPI)
   - Decide on JPEG quality tuning (85% recommended start)

### For Shiherlis (Testing Strategy)

1. **Tier 1 tests (straightforward):**
   - Unit tests for categorization logic (mock ZIP entries, verify counts)
   - Edge cases: empty presentations, presentations with no media, minimal masters
   - Test with 3+ real presentations (small, medium, complex)

2. **Tier 2 tests (rigorous required):**
   - **#83 (Removal):**
     - Prevent orphan removal (mock scenarios where removing master would break layouts)
     - Validate before/after with OpenXmlValidator
     - E2E: Create presentation with unused masters, remove, open in PowerPoint
   - **#84 (Dedup):**
     - Hash collision tests (confirm same content hashes correctly)
     - Reference integrity (all image references point to canonical after dedup)
     - E2E: Create presentation with 3+ duplicate images across slides, dedup, open in PowerPoint, verify visual correctness
   - **#85 (Image Compression):**
     - Downscaling logic (images larger than display bounds are downscaled, small images untouched)
     - Format conversion (BMP→PNG, etc.)
     - Quality trade-offs (verify JPEG quality setting produces acceptable results)
     - E2E: Compress presentation with large/mixed-format images, open in PowerPoint, verify size reduction + acceptable quality

3. **Cross-tool baseline:**
   - After each tool, re-run Phase 1–3 regression tests (ensure no breakage in existing functionality)
   - All Phase 4 tools must have 3+ unit test cases + comprehensive E2E validation

---

## Codebase Readiness Assessment

### Current State
- **Build Health:** Passing, 80 build warnings (acceptable)
- **Test Baseline:** 377 tests passing, 0 failures
- **Existing Patterns We'll Leverage:**
  - `PresentationService` (partial classes for organized tool methods) ✅
  - OpenXML traversal (slide/layout/master enumeration already done) ✅
  - MCP tool patterns (`[McpServerToolType]`, `[McpServerTool]`, XML doc comments) ✅
  - Error handling + structured JSON responses ✅

### New Capabilities Needed for Phase 4
1. **ZIP enumeration** (System.IO.Compression) — new
2. **SHA256 hashing** (System.Security.Cryptography) — new
3. **Image processing** (SkiaSharp) — new dependency
4. **Relationship mutation** (removing parts, updating references) — new, requires careful testing
5. **OpenXmlValidator** usage — likely already available but needs integration

### No Blocker Issues Identified
- Existing architecture supports both analysis and mutation tools
- Dependencies are lightweight and standard
- Test infrastructure (xUnit v3 on MTP) is ready

---

## Squad Assignments & Handoff

### Current State (GitHub Issues)
- All 7 issues currently assigned to `squad:shiherlis`
- Some also tagged `squad:copilot` (for documentation)

### Recommended Reassignment (Proper Squad Fit)

| Issue | Type | Current Owner | Recommended Owner | Reason |
|-------|------|---------------|-------------------|--------|
| #80 (P4-1) | Analysis/Tool | Shiherlis | **Cheritto** | Backend dev → tool implementation |
| #81 (P4-2) | Analysis/Tool | Shiherlis | **Cheritto** | Backend dev → tool implementation |
| #82 (P4-3) | Analysis/Tool | Shiherlis | **Cheritto** | Backend dev → tool implementation |
| #83 (P4-4) | Optimization/Tool | Shiherlis | **Cheritto** | Backend dev → tool implementation (high risk; needs testing support) |
| #84 (P4-5) | Optimization/Tool | Shiherlis | **Cheritto** | Backend dev → tool implementation (high risk; needs testing support) |
| #85 (P4-6) | Optimization/Tool | Shiherlis | **Cheritto** | Backend dev → tool implementation + SkiaSharp research |
| #86 (P4-7) | Analysis/Deferred | Shiherlis | **Parked** | Mark `go:no`, reassign to backlog epic (future spike) |

### Testing Assignments
- Create separate testing issues for each Tier:
  - **T1-Test** (Shiherlis): Unit + E2E for #80, #81, #82 (2–3 hours)
  - **T2-Test** (Shiherlis): Unit + E2E + PowerPoint round-trip for #83, #84, #85 (6–8 hours)
  - These testing issues should have hard dependencies on corresponding tool issues

### Documentation Assignments
- Create **Phase 4 Documentation** issue (@copilot):
  - Add 3 new tools to TOOL_REFERENCE.md (#80, #81, #82 analysis tools)
  - Add 3 new tools to TOOL_REFERENCE.md (#83, #84, #85 optimization tools)
  - Add example workflow: "How to compress a PowerPoint file" (combining multiple tools)
  - Note: Video optimization tool (#86) not documented (parked)

---

## Risk Assessment

### Technical Risks

| Risk | Likelihood | Impact | Mitigation |
|------|-----------|--------|-----------|
| Relationship breakage (removing parts orphans refs) | **MEDIUM** | **HIGH** (file corruption) | Comprehensive validation before/after; E2E PowerPoint round-trip mandatory |
| Image quality degradation (compression too aggressive) | **MEDIUM** | **MEDIUM** (user complaint) | Document quality settings; E2E visual inspection; default to conservative (85% JPEG) |
| SkiaSharp cross-platform issues | **LOW** | **MEDIUM** (blocks builds on some platforms) | Dependency research required (Nate); documented in README |
| OpenXmlValidator insufficiency | **MEDIUM** | **MEDIUM** (files pass validator but fail in PowerPoint) | Always test round-trip: save→open in PowerPoint; don't trust validator alone |
| ZIP enumeration missing parts | **LOW** | **LOW** (incorrect size reporting) | Unit tests with mock ZIP; comprehensive test presentations |

### Schedule Risks

| Risk | Likelihood | Impact | Mitigation |
|------|-----------|--------|-----------|
| Tier 2 PowerPoint testing takes longer | **MEDIUM** | **MEDIUM** (schedule slip) | Allocate extra time; test early and often; have test presentations ready |
| SkiaSharp dependency complication | **LOW** | **MEDIUM** (blocks P4-6) | Research early (Nate task); decide library early |
| Regression in Phase 1–3 functionality | **LOW** | **HIGH** (discovery during Phase 4 testing) | Run full baseline test suite before starting each tier |

### PowerPoint Compatibility Risks

**Critical Assumption:** OpenXML compliance ≠ PowerPoint compatibility.

- **Files passing OpenXmlValidator can still fail to open in PowerPoint** if:
  - Relationship structure is corrupted (dangling references, cycles)
  - Required parts are missing (orphaned layouts, images)
  - Metadata is inconsistent (slide IDs, part IDs)

**Mitigation:**
- All Tier 2 operations must include E2E PowerPoint round-trip tests
- Test on Windows PowerPoint (primary) + consider Mac/Office Online (if resources allow)
- Have a set of known-good test presentations; verify no regression

---

## Success Criteria

### Tier 1 (Analysis Tools) — Clear `go:needs-research`
- [x] All 3 issues (#80, #81, #82) have clear acceptance criteria
- [x] No blocking technical unknowns
- [ ] Nate completes OpenXML pattern research (master/layout, media refs, validation strategy)
- [x] Unit tests scoped: 3+ cases per tool
- [x] E2E tests scoped: run on 3+ real presentations

### Tier 2 (Optimization Tools) — Ready for Implementation
- [x] All 3 issues (#83, #84, #85) have clear acceptance criteria + dependencies
- [ ] Nate completes relationship semantics + SkiaSharp trade-off research
- [ ] Cheritto starts with P4-3 analysis logic, then moves to P4-4 removal
- [x] Testing strategy includes PowerPoint round-trip for each tool
- [x] Risk mitigation documented (validation, error handling, logging)

### Overall Phase 4 Success
- All Tier 1 + Tier 2 issues closed
- 400+ tests passing (no regression from Phase 1–3 baseline of 377)
- All Tier 2 E2E tests include successful PowerPoint round-trip
- Documentation updated (README + TOOL_REFERENCE.md)
- No Phase 1–3 functionality broken

---

## Known Unknowns & Future Work

1. **Video optimization** (#86) — parked for future spike; may require ffmpeg or external API
2. **Cross-platform testing** — Phase 4 assumes Windows PowerPoint (primary); Mac/Office Online untested
3. **Performance tuning** — If presentations are very large (1GB+), ZIP enumeration and hashing may be slow; could optimize later
4. **Presentation-level size reduction** — Future phase might include removing unused slide themes, consolidating color schemes, etc. (out of scope for Phase 4)

---

## Next Steps

1. **Nate:** Research OpenXML patterns (master/layout semantics, media reference structure, SkiaSharp trade-offs, validation strategy) — 2–3 hours, due before Cheritto starts
2. **Cheritto:** Implement Tier 1 tools (#80, #81, #82) in sequence; aim for mid-week completion
3. **Shiherlis:** Prepare test scenarios (real presentations, edge cases); stand up testing infrastructure for E2E PowerPoint validation
4. **@Copilot:** Draft Phase 4 documentation issue; add placeholders for new tools (fill in once implementations complete)
5. **McCauley:** Clear `go:needs-research` on Tier 1 issues once Nate research complete; approve issue reassignments; monitor PowerPoint compatibility risk

---

## Appendix: Tier Definitions

**Tier 1 — Read-Only Analysis:**
- No PPTX mutations
- Low implementation complexity
- No PowerPoint compatibility risk
- Provides data foundation for cleanup

**Tier 2 — Write Operations:**
- Mutates PPTX package (removes/updates parts)
- Requires strict validation (OpenXmlValidator + PowerPoint round-trip)
- Higher implementation complexity
- Moderate schedule risk (testing is time-intensive)

**Tier 3 — Deferred/Future:**
- Low priority; valuable later
- Parking lot for future spike or community request
- Not starting Phase 4

---

**Prepared by:** McCauley (Lead)  
**For approval by:** Jon Galloway  
**Status:** Ready for Squad Handoff


---

### Phase 4: OpenXML Patterns Research (2026-03-24 Nate)

**Researcher:** Nate (Consulting Dev)  
**Status:** Complete  

#### Summary

Completed feasibility analysis for Phase 4 issues—all 7 are highly viable. Established OpenXML API patterns and dependency recommendations.

#### Phase 4: OpenXML Patterns Research for Presentation Optimization & Media Analysis

**Research Date:** 2026-03-24  
**Researcher:** Nate (Consulting Dev)  
**Task:** Feasibility analysis for Phase 4 issues (#80–#86) — file size breakdown, media analysis, layout deduplication, optimization patterns.

**Research Scope:** Examined MarpToPptx (prior art for media/extraction), dotnet-mcp (MCP patterns), current pptx-mcp PresentationService, and OpenXML SDK v3.3.0 capabilities.

---

## Executive Summary

**✅ Highly Feasible (High Confidence):**
- **#80 — File size breakdown:** Direct ZIP access available; PackagePart enumeration mature
- **#81 — Media asset analysis:** Prior art in MarpToPptx extraction; ImagePart/MediaDataPart APIs proven
- **#82 — Unused layout detection:** Relationship traversal patterns established; straightforward cross-reference logic

**⚠️ Medium Feasibility (Medium Confidence):**
- **#83 — Remove unused layouts:** Safe if relationship cleanup done carefully; no PowerPoint round-trip risk
- **#84 — Deduplicate media:** Hash-based comparison works; relationship redirection is standard OPC, but needs edge-case testing

**🤔 Specialized Feasibility:**
- **#85 — Image compression:** No built-in SDK support; requires external image library or shell integration (e.g., ImageSharp, System.Drawing)
- **#86 — Video optimization (analysis-only):** Viable for metadata extraction, but codec detection may need external library (e.g., MediaInfo CLI)

---

## Phase 4 Issues: Detailed Feasibility Analysis

### Issue #80: Analyze File Size Breakdown

**Problem:** Enumerate ZIP parts in a PPTX, categorize by content type (slides, images, themes, etc.), measure sizes.

**Feasibility:** ✅ **HIGH (High Confidence)**

**Prior Art:**
- MarpToPptx: `System.IO.Compression.ZipArchive` used directly in `OpenXmlPptxRenderer.cs`
- No external dependencies; standard .NET API

**OpenXML API Surface:**
```csharp
// Two approaches:

// 1. Via PresentationDocument + PackagePart enumeration (RECOMMENDED)
using var doc = PresentationDocument.Open(filePath, false);
var presentationPart = doc.PresentationPart;

// Access underlying Package (OPC container)
var package = presentationPart.OpenXmlPackage.Package;

// Enumerate all parts by type
foreach (var part in package.GetParts())
{
    var partUri = part.Uri;           // e.g., /ppt/slides/slide1.xml
    var contentType = part.ContentType; // e.g., application/vnd.openxmlformats-officedocument.presentationml.slide+xml
    var size = part.GetStream().Length; // Byte size
}

// 2. Via direct ZipArchive (RAW ZIP access, if needed for compression metadata)
using var zip = ZipFile.OpenRead(filePath);
foreach (var entry in zip.Entries)
{
    var path = entry.FullName;
    var compressedSize = entry.CompressedLength;
    var uncompressedSize = entry.Length;
    var compressionRatio = (double)entry.CompressedLength / entry.Length;
}
```

**Categorization Logic:**
```
Slides:       /ppt/slides/slide*.xml
Media:        /ppt/media/* (images, audio, video)
Themes:       /ppt/theme/* 
Layouts:      /ppt/slideLayouts/*
Masters:      /ppt/slideMasters/*
Relationships:/ppt/slides/_rels/slide*.xml.rels
Other:        /docProps/*, /ppt/_rels/*, [Content_Types].xml
```

**Gotchas:**
- `PackagePart` and `ZipArchive` give different compression info; ZipArchive shows actual disk savings, PackagePart shows logical size
- Some PKG parts may be excluded from ZIP (edge case; rare in PPTX)
- Relationship files count separately but are often highly compressible

**Dependencies:** None (standard .NET 10 APIs)

**Code Sketch:**
```csharp
public class FileSizeBreakdown
{
    public long SlideSize { get; set; }
    public long MediaSize { get; set; }
    public long ThemeSize { get; set; }
    public long LayoutSize { get; set; }
    public long MasterSize { get; set; }
    public long RelationshipSize { get; set; }
    public long OtherSize { get; set; }
}

public FileSizeBreakdown AnalyzePptxSize(string filePath)
{
    using var doc = PresentationDocument.Open(filePath, false);
    var package = doc.PresentationPart.OpenXmlPackage.Package;
    var breakdown = new FileSizeBreakdown();

    foreach (var part in package.GetParts())
    {
        var uri = part.Uri.ToString().ToLower();
        var size = part.GetStream().Length;

        if (uri.Contains("/slides/slide")) breakdown.SlideSize += size;
        else if (uri.Contains("/media/")) breakdown.MediaSize += size;
        else if (uri.Contains("/theme/")) breakdown.ThemeSize += size;
        else if (uri.Contains("/slideLayouts/")) breakdown.LayoutSize += size;
        else if (uri.Contains("/slideMasters/")) breakdown.MasterSize += size;
        else if (uri.EndsWith(".rels")) breakdown.RelationshipSize += size;
        else breakdown.OtherSize += size;
    }

    return breakdown;
}
```

**Validation:** Returns accurate JSON report; PowerPoint file remains unchanged (read-only access)

---

### Issue #81: List and Analyze Media Assets

**Problem:** Find all images, audio, video in PPTX. Report: type, dimensions (images), format, compression info.

**Feasibility:** ✅ **HIGH (High Confidence)**

**Prior Art:**
- MarpToPptx: `PptxMarkdownExporter.Media.cs` — extracts images using `ImagePart`, `MediaDataPart`; computes SHA256 hashes for deduplication
- Pattern: Enumerate `Picture` shapes → resolve `Blip.Embed` relationship → get `ImagePart` → stream metadata

**OpenXML API Surface:**
```csharp
// Image extraction
using var doc = PresentationDocument.Open(filePath, false);
foreach (var slidePart in doc.PresentationPart.SlideParts)
{
    var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
    foreach (var picture in shapeTree.Elements<Picture>())
    {
        var relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
        if (relationshipId != null && slidePart.TryGetPartById(relationshipId, out var part))
        {
            if (part is ImagePart imagePart)
            {
                var contentType = imagePart.ContentType; // "image/png", "image/jpeg", etc.
                using var stream = imagePart.GetStream();
                // Extract dimensions, file size, etc.
            }
        }
    }
}

// Video/Audio extraction (MediaDataPart)
// NOTE: MediaDataParts are references, not direct properties
// Access via Blip.VideoFromFile or AudioFromFile child elements
foreach (var blip in shapeTree.Descendants<Blip>())
{
    var video = blip.GetFirstChild<VideoFromFile>();
    if (video != null)
    {
        var embedRelId = video.Embed?.Value; // Relationship to MediaDataPart
        if (slidePart.TryGetPartById(embedRelId, out var part) && part is MediaDataPart mediaPart)
        {
            var format = mediaPart.ContentType; // e.g., "video/mp4"
            var size = mediaPart.GetStream().Length;
        }
    }
}
```

**Image Dimension Extraction:**
```csharp
// For PNG/JPEG, parse header to get dimensions
// MarpToPptx uses similar approach in diagnostics
private (int Width, int Height) GetImageDimensions(ImagePart imagePart)
{
    using var stream = imagePart.GetStream();
    var buffer = new byte[8];
    stream.Read(buffer, 0, 8);

    if (imagePart.ContentType == "image/png")
    {
        // PNG: width/height at bytes 16-24 (big-endian)
        stream.Seek(16, SeekOrigin.Begin);
        var widthBytes = new byte[4];
        stream.Read(widthBytes, 0, 4);
        var width = BitConverter.ToInt32(widthBytes.Reverse().ToArray(), 0);
        // ... similar for height
    }
    else if (imagePart.ContentType.Contains("jpeg"))
    {
        // JPEG: SOF marker parsing (more complex; use external library preferred)
    }

    return (0, 0); // Fallback
}
```

**Gotchas:**
- Video/Audio parts may not always have direct relationships; check `VideoFromFile`/`AudioFromFile` wrapper elements
- Image dimensions require parsing file headers (PNG/JPEG); no direct API
- Some media may be embedded in charts (ChartPart) — separate enumeration needed
- Compression ratio (in-file vs. uncompressed) available via ZipArchive but not PackagePart

**Dependencies:** None required for basic analysis; recommend `ImageSharp` (optional) for robust dimension parsing

**Code Sketch:**
```csharp
public class MediaAnalysis
{
    public string Type { get; set; } // "image", "video", "audio"
    public string Format { get; set; } // "png", "jpeg", "mp4", etc.
    public long Size { get; set; }
    public int? Width { get; set; } // Images only
    public int? Height { get; set; } // Images only
    public string ContentType { get; set; } // MIME type
    public string Hash { get; set; } // SHA256 hex
}

public List<MediaAnalysis> AnalyzeMedia(string filePath)
{
    var result = new List<MediaAnalysis>();
    using var doc = PresentationDocument.Open(filePath, false);
    
    foreach (var slidePart in doc.PresentationPart.SlideParts)
    {
        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        
        // Images
        foreach (var picture in shapeTree.Elements<Picture>())
        {
            var relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
            if (relationshipId != null && slidePart.TryGetPartById(relationshipId, out var part))
            {
                if (part is ImagePart imagePart)
                {
                    using var stream = imagePart.GetStream();
                    var hash = Convert.ToHexString(SHA256.HashData(stream));
                    var analysis = new MediaAnalysis
                    {
                        Type = "image",
                        Format = GetFormatFromContentType(imagePart.ContentType),
                        Size = stream.Length,
                        ContentType = imagePart.ContentType,
                        Hash = hash
                    };
                    result.Add(analysis);
                }
            }
        }
    }

    return result;
}
```

**Validation:** JSON array of media objects; no file modification

---

### Issue #82: Find Unused Slide Masters and Layouts

**Problem:** Traverse PresentationPart → SlideMasterParts → SlideLayoutParts. Cross-reference against actual slide usage. Report: which layouts/masters are unused.

**Feasibility:** ✅ **HIGH (High Confidence)**

**Prior Art:**
- pptx-mcp: `GetLayouts()` method already enumerates masters and layouts
- MarpToPptx: Template selection logic (`SlideTemplateSelector`) resolves layout relationships
- Standard OpenXML traversal pattern

**OpenXML API Surface:**
```csharp
// 1. Enumerate all layouts in presentation
var allLayouts = new Dictionary<string, SlideLayoutPart>();
foreach (var masterPart in presentationPart.SlideMasterParts)
{
    foreach (var layoutPart in masterPart.SlideLayoutParts)
    {
        var layoutId = presentationPart.GetIdOfPart(layoutPart); // Relationship ID
        allLayouts[layoutId] = layoutPart;
    }
}

// 2. Find layouts actually used in slides
var usedLayouts = new HashSet<string>();
foreach (var slidePart in presentationPart.SlideParts)
{
    var layoutId = slidePart.Slide.CommonSlideData?.SlideLayoutId; // NOT directly available!
    // ACTUAL APPROACH: Get SlideLayoutPart directly
    var slideLayoutPart = slidePart.SlideLayoutPart;
    if (slideLayoutPart != null)
    {
        var layoutId = presentationPart.GetIdOfPart(slideLayoutPart);
        usedLayouts.Add(layoutId);
    }
}

// 3. Find unused
var unused = allLayouts.Keys.Except(usedLayouts).ToList();
```

**Gotcha — Tricky Relationship Resolution:**
- Slides have implicit relationship to their SlideLayoutPart (via `slidePart.SlideLayoutPart`)
- This relationship is stored in slide's `.rels` file, not in the slide XML itself
- To get the relationship ID: `slidePart.GetIdOfPart(slidePart.SlideLayoutPart)`

**Master Usage:**
- All layouts belong to a master; if a layout is unused, its master may also be unused
- However, a master may be referenced in metadata or for theme/style inheritance → safer to report "layout-unused" rather than "master-unused"

**Gotchas:**
- Blank layouts are sometimes kept for reusability; cross-check before suggesting deletion
- Layouts may have custom names (accessible via `layoutPart.SlideLayout.CommonSlideData?.Name?.Value`)
- Some layouts may be referenced by Name in Marp metadata (MarpToPptx pattern) but not actually used

**Dependencies:** None

**Code Sketch:**
```csharp
public class LayoutUsageReport
{
    public string LayoutName { get; set; }
    public string MasterName { get; set; }
    public int SlideCount { get; set; } // Slides using this layout
    public bool IsUsed { get; set; }
}

public List<LayoutUsageReport> FindUnusedLayouts(string filePath)
{
    using var doc = PresentationDocument.Open(filePath, false);
    var presentationPart = doc.PresentationPart;
    
    // Build inventory of all layouts
    var allLayouts = new Dictionary<string, (SlideLayoutPart, SlideMasterPart)>();
    foreach (var masterPart in presentationPart.SlideMasterParts)
    {
        foreach (var layoutPart in masterPart.SlideLayoutParts)
        {
            var layoutId = presentationPart.GetIdOfPart(layoutPart);
            allLayouts[layoutId] = (layoutPart, masterPart);
        }
    }

    // Count usage
    var usageCounts = new Dictionary<string, int>();
    foreach (var slidePart in presentationPart.SlideParts)
    {
        if (slidePart.SlideLayoutPart != null)
        {
            var layoutId = presentationPart.GetIdOfPart(slidePart.SlideLayoutPart);
            usageCounts[layoutId] = usageCounts.TryGetValue(layoutId, out var count) ? count + 1 : 1;
        }
    }

    // Report
    var report = new List<LayoutUsageReport>();
    foreach (var (layoutId, (layoutPart, masterPart)) in allLayouts)
    {
        var layoutName = layoutPart.SlideLayout.CommonSlideData?.Name?.Value ?? "Unnamed";
        var masterName = masterPart.SlideMaster.CommonSlideData?.Name?.Value ?? "Unnamed";
        var count = usageCounts.TryGetValue(layoutId, out var c) ? c : 0;

        report.Add(new LayoutUsageReport
        {
            LayoutName = layoutName,
            MasterName = masterName,
            SlideCount = count,
            IsUsed = count > 0
        });
    }

    return report;
}
```

**Validation:** JSON array; no file modification

---

### Issue #83: Remove Unused Slide Masters and Layouts

**Problem:** Delete unused layouts from issue #82. Safe cleanup: relationship integrity, content type cleanup.

**Feasibility:** ⚠️ **MEDIUM (Medium Confidence)**

**Risk Level:** MEDIUM — Relationship cleanup is robust in OpenXML SDK, but PowerPoint round-trip testing essential.

**OpenXML API Surface:**
```csharp
// Safe deletion pattern (from OpenXML SDK docs):
1. Identify unused layout (from issue #82)
2. Get SlideLayoutPart reference
3. Call presentationPart.DeletePart(layoutPart)

// SDK automatically handles:
// - Removing relationship from master
// - Updating [Content_Types].xml
// - Removing layout's own sub-parts (relationships)

using var doc = PresentationDocument.Open(filePath, true);
var presentationPart = doc.PresentationPart;

foreach (var layoutId in unusedLayoutIds)
{
    var layoutPart = presentationPart.GetPartById(layoutId) as SlideLayoutPart;
    if (layoutPart != null)
    {
        presentationPart.DeletePart(layoutPart); // Master is still referenced; SDK keeps it
    }
}

doc.PresentationPart.Presentation.Save();
```

**Master Deletion (More Complex):**
```csharp
// Only safe if:
// 1. Master has no layouts (or all deleted already)
// 2. No slides reference master's theme
// 3. No XML metadata references it

var masterPart = presentationPart.GetPartById(masterId) as SlideMasterPart;
if (masterPart != null && masterPart.SlideLayoutParts.Count == 0)
{
    presentationPart.DeletePart(masterPart); // Also removes from SlideMasterIdList
}
```

**Gotchas:**
- **PowerPoint round-trip risk:** Deleting layouts sometimes causes "missing template" warnings in PowerPoint. Always test in PowerPoint after deletion.
- **Theme inheritance:** If master is deleted but theme is still referenced elsewhere → corruption risk
- **Relationships cleanup:** OpenXML SDK is robust, but edge cases exist (custom XML, VBA macros referencing layouts)
- **Atomic save:** If deletion fails mid-operation, PPTX may be corrupted. Wrap in try/finally.

**Prior Art:**
- MarpToPptx: `DeleteSlideParts()` method in `OpenXmlPptxRenderer.cs` handles part deletion safely
- dotnet-mcp: No direct prior art, but general pattern: SDK handles relationship cleanup if called correctly

**Dependencies:** None; DocumentFormat.OpenXml 3.3.0 sufficient

**Code Sketch:**
```csharp
public class RemovalResult
{
    public bool Success { get; set; }
    public int LayoutsRemoved { get; set; }
    public int MastersRemoved { get; set; }
    public string Message { get; set; }
}

public RemovalResult RemoveUnusedLayouts(string filePath, List<string> unusedLayoutIds)
{
    using var doc = PresentationDocument.Open(filePath, true);
    var presentationPart = doc.PresentationPart;
    var result = new RemovalResult { Success = true };

    try
    {
        foreach (var layoutId in unusedLayoutIds)
        {
            if (presentationPart.TryGetPartById(layoutId, out var part) && part is SlideLayoutPart layoutPart)
            {
                presentationPart.DeletePart(layoutPart);
                result.LayoutsRemoved++;
            }
        }

        // Optionally remove orphaned masters
        foreach (var masterPart in presentationPart.SlideMasterParts.ToList())
        {
            if (masterPart.SlideLayoutParts.Count == 0)
            {
                presentationPart.DeletePart(masterPart);
                result.MastersRemoved++;
            }
        }

        doc.Save();
    }
    catch (Exception ex)
    {
        result.Success = false;
        result.Message = ex.Message;
    }

    return result;
}
```

**Validation:** 
1. File size reduced
2. Open in PowerPoint without warnings
3. Re-scan with issue #82 to confirm removal

---

### Issue #84: Deduplicate Identical Media

**Problem:** Find duplicate media blobs (hash comparison). Redirect relationships from multiple slides to single media part.

**Feasibility:** ⚠️ **MEDIUM-HIGH (Medium Confidence)**

**Prior Art:**
- MarpToPptx: `ComputeImageSignature()` in `PptxMarkdownExporter.Media.cs` — uses SHA256 hash on stream
- Pattern: Hash all images → group by hash → redirect relationships

**OpenXML API Surface:**
```csharp
// 1. Enumerate all media and compute hashes
var mediaByHash = new Dictionary<string, ImagePart>();
foreach (var slidePart in presentationPart.SlideParts)
{
    var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
    foreach (var picture in shapeTree.Elements<Picture>())
    {
        var relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
        if (relationshipId != null && slidePart.TryGetPartById(relationshipId, out var part) && part is ImagePart imagePart)
        {
            using var stream = imagePart.GetStream();
            var hash = Convert.ToHexString(SHA256.HashData(stream));
            
            if (!mediaByHash.ContainsKey(hash))
                mediaByHash[hash] = imagePart;
            // else: duplicate found
        }
    }
}

// 2. Redirect relationships to canonical part
var duplicateImages = new List<ImagePart>();
foreach (var group in groupedByHash.Skip(1)) // Skip first (canonical)
{
    foreach (var imagePart in group.Value.Skip(1))
    {
        duplicateImages.Add(imagePart);
    }
}

// 3. Update Blip.Embed to point to canonical
// This is TRICKY: you must update the relationship, not delete the part yet
foreach (var slidePart in presentationPart.SlideParts)
{
    foreach (var picture in shapeTree.Elements<Picture>())
    {
        var relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
        var imagePart = slidePart.GetPartById(relationshipId) as ImagePart;
        var hash = ComputeHash(imagePart);
        
        if (mediaByHash[hash] != imagePart) // It's a duplicate
        {
            // Remove old relationship
            slidePart.DeleteRelationship(relationshipId);
            
            // Add new relationship to canonical
            var newRelId = slidePart.CreateRelationshipToOtherPart(mediaByHash[hash]);
            picture.BlipFill.Blip.Embed.Value = newRelId;
        }
    }
}

// 4. Delete orphaned parts
foreach (var orphan in duplicateImages)
{
    presentationPart.DeletePart(orphan);
}

doc.Save();
```

**Gotchas:**
- **Relationship redirection is manual:** SDK doesn't have `RedirectRelationship()` helper; you must delete+recreate
- **Blip.Embed is StringValue:** Setting it to new relationship ID requires careful XML manipulation
- **File format edge cases:** Two visually identical images may differ in metadata (EXIF, compression); hash detects true binary duplicates only
- **Charts, SmartArt:** May contain embedded media not directly via Picture shapes; need separate enumeration
- **Atomic transaction:** If redirect fails mid-operation, relationships break. Wrap in backup pattern.

**Dependencies:** None for hashing; optional `ImageSharp` for perceptual hashing (fuzzy deduplication)

**Code Sketch:**
```csharp
public class DeduplicationResult
{
    public bool Success { get; set; }
    public int DuplicatesFound { get; set; }
    public int DuplicatesRemoved { get; set; }
    public long BytesSaved { get; set; }
    public string Message { get; set; }
}

public DeduplicationResult DeduplicateMedia(string filePath)
{
    using var doc = PresentationDocument.Open(filePath, true);
    var presentationPart = doc.PresentationPart;
    var result = new DeduplicationResult { Success = true };

    try
    {
        // Build hash inventory
        var mediaByHash = new Dictionary<string, List<(SlidePart, string, ImagePart)>>();
        
        foreach (var slidePart in presentationPart.SlideParts)
        {
            var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
            foreach (var picture in shapeTree.Elements<Picture>())
            {
                var relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
                if (relationshipId != null && slidePart.TryGetPartById(relationshipId, out var part) && part is ImagePart imagePart)
                {
                    using var stream = imagePart.GetStream();
                    var hash = Convert.ToHexString(SHA256.HashData(stream));
                    
                    if (!mediaByHash.ContainsKey(hash))
                        mediaByHash[hash] = new();
                    
                    mediaByHash[hash].Add((slidePart, relationshipId, imagePart));
                }
            }
        }

        // Redirect and delete duplicates
        var orphans = new HashSet<ImagePart>();
        foreach (var (hash, instances) in mediaByHash)
        {
            if (instances.Count > 1)
            {
                result.DuplicatesFound += instances.Count - 1;
                var canonical = instances[0].Item3;

                foreach (var (slidePart, oldRelId, duplicateImage) in instances.Skip(1))
                {
                    // Delete old relationship
                    slidePart.DeleteRelationship(oldRelId);
                    
                    // Create new relationship to canonical
                    var newRelId = slidePart.CreateRelationshipToOtherPart(canonical);
                    
                    // Update Blip.Embed
                    var picture = slidePart.Slide.CommonSlideData.ShapeTree
                        .Elements<Picture>()
                        .FirstOrDefault(p => p.BlipFill?.Blip?.Embed?.Value == oldRelId); // Need to track better
                    if (picture != null)
                        picture.BlipFill.Blip.Embed.Value = newRelId;
                    
                    orphans.Add(duplicateImage);
                }
            }
        }

        // Delete orphans
        foreach (var orphan in orphans)
        {
            result.BytesSaved += orphan.GetStream().Length;
            presentationPart.DeletePart(orphan);
            result.DuplicatesRemoved++;
        }

        doc.Save();
    }
    catch (Exception ex)
    {
        result.Success = false;
        result.Message = ex.Message;
    }

    return result;
}
```

**Validation:**
1. File opens in PowerPoint without corruption
2. Visual verification: images still displayed correctly
3. File size reduced by deduplication count × average image size
4. Re-scan with issue #81 to confirm dedup count

---

### Issue #85: Compress/Optimize Images

**Problem:** Re-encode images at lower quality/resolution. No heavy dependencies.

**Feasibility:** ⚠️ **MEDIUM (Medium Confidence)**

**Complexity:** MEDIUM-HIGH — requires image processing library or CLI integration

**Options:**

**Option A: System.Drawing (Windows-only, deprecated)**
```csharp
// NOT RECOMMENDED for .NET 5+; marked obsolete
// Only available on Windows; brings Platform.Windows dependency
```

**Option B: ImageSharp (NuGet, recommended)**
```csharp
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;

public void CompressImage(ImagePart imagePart, int quality = 85, int? maxWidth = 1920)
{
    using var stream = imagePart.GetStream(FileMode.Open, FileAccess.ReadWrite);
    
    using var image = Image.Load(stream);
    
    // Resize if needed
    if (maxWidth.HasValue && image.Width > maxWidth.Value)
    {
        var height = (int)(image.Height * ((double)maxWidth.Value / image.Width));
        image.Mutate(x => x.Resize(maxWidth.Value, height));
    }
    
    // Re-encode at lower quality
    var options = new JpegEncoder { Quality = quality };
    stream.Seek(0, SeekOrigin.Begin);
    stream.SetLength(0);
    image.SaveAsJpeg(stream, options);
}
```

**Option C: ImageMagick CLI (external process)**
```csharp
// Shell out to ImageMagick (requires CLI installed)
using var process = new System.Diagnostics.Process
{
    StartInfo = new System.Diagnostics.ProcessStartInfo
    {
        FileName = "magick",
        Arguments = $"input.jpg -quality 85 -resize 1920x1080 output.jpg",
        UseShellExecute = false,
        RedirectStandardOutput = true
    }
};
process.Start();
process.WaitForExit();
```

**Option D: SkiaSharp (cross-platform, high-performance)**
```csharp
using SkiaSharp;

public void CompressImageSkia(ImagePart imagePart, int quality = 85)
{
    using var stream = imagePart.GetStream(FileMode.Open, FileAccess.ReadWrite);
    
    using var skBitmap = SKBitmap.Decode(stream);
    using var image = SKImage.FromBitmap(skBitmap);
    using var encoded = image.Encode(SKEncodedImageFormat.Jpeg, quality);
    
    stream.Seek(0, SeekOrigin.Begin);
    stream.SetLength(0);
    encoded.SaveTo(stream);
}
```

**Gotchas:**
- **Lossy vs. Lossless:** JPEG compression is lossy; quality 85 is sweet spot (visually indistinguishable, ~30% file savings)
- **PNG optimization:** PNGs are lossless; compression is limited to ZIP deflate tuning; ImageSharp can re-encode at higher deflate level but savings are ~10–20%
- **Format conversion:** Converting PNG→JPEG may change appearance (no alpha channel); risky without explicit user opt-in
- **Metadata loss:** Re-encoding strips EXIF, color profiles; may impact accessibility
- **Performance:** Processing large images (4K+) is slow; consider progress reporting
- **PowerPoint compatibility:** Some codecs (HEIC, WEBP) may not display in PowerPoint; stick to PNG/JPEG

**Dependencies:**
- **ImageSharp (SixLabors):** NuGet `SixLabors.ImageSharp` (~3 MB)
- **SkiaSharp:** NuGet `SkiaSharp` (~15 MB, includes native binaries)
- Neither is "heavy" by modern standards; ImageSharp preferred for .NET-only stack

**Recommendation:**
- **Add `SixLabors.ImageSharp` as optional dependency** (with feature flag or configuration)
- Document that compression is lossy; user must opt-in
- Provide presets: `Light` (quality 95, no resize), `Medium` (quality 85, max 1920px), `Aggressive` (quality 75, max 1280px)

**Code Sketch:**
```csharp
public class ImageCompressionResult
{
    public bool Success { get; set; }
    public long OriginalSize { get; set; }
    public long CompressedSize { get; set; }
    public double CompressionRatio { get; set; }
    public string Message { get; set; }
}

public ImageCompressionResult CompressImage(string filePath, int slideNumber, string imageName, int quality = 85, int? maxWidth = 1920)
{
    using var doc = PresentationDocument.Open(filePath, true);
    var slidePart = GetSlidePart(doc, slideNumber - 1);
    
    var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
    var picture = shapeTree.Elements<Picture>()
        .FirstOrDefault(p => /* match by name */);
    
    if (picture?.BlipFill?.Blip?.Embed?.Value is null)
        return new() { Success = false, Message = "Image not found" };
    
    if (slidePart.TryGetPartById(picture.BlipFill.Blip.Embed.Value, out var part) && part is ImagePart imagePart)
    {
        try
        {
            using (var stream = imagePart.GetStream(FileMode.Open, FileAccess.ReadWrite))
            {
                var originalSize = stream.Length;
                
                using (var image = SixLabors.ImageSharp.Image.Load(stream))
                {
                    if (maxWidth.HasValue && image.Width > maxWidth.Value)
                    {
                        var height = (int)(image.Height * ((double)maxWidth.Value / image.Width));
                        image.Mutate(x => x.Resize(maxWidth.Value, height));
                    }
                    
                    var options = new JpegEncoder { Quality = quality };
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.SetLength(0);
                    image.SaveAsJpeg(stream, options);
                }
                
                var compressedSize = stream.Length;
                doc.Save();
                
                return new()
                {
                    Success = true,
                    OriginalSize = originalSize,
                    CompressedSize = compressedSize,
                    CompressionRatio = (double)compressedSize / originalSize,
                    Message = $"Compressed from {originalSize} to {compressedSize} bytes"
                };
            }
        }
        catch (Exception ex)
        {
            return new() { Success = false, Message = ex.Message };
        }
    }
    
    return new() { Success = false, Message = "ImagePart not found" };
}
```

**Validation:**
1. Image dimensions correct (or resized as specified)
2. File opens in PowerPoint without corruption
3. Visual quality acceptable for specified compression level
4. File size reduced by expected ratio

---

### Issue #86: Video Optimization (Analysis Only)

**Problem:** Why is this `go:no`? What would analysis-only look like?

**Analysis:** This issue is marked `go:no` (lowest priority / out of scope) because:
1. **OpenXML SDK video support is minimal** — no direct codec introspection APIs
2. **External tooling required** — MediaInfo or FFProbe CLI needed for codec/bitrate metadata
3. **Use case unclear** — "optimize" is vague (re-encode? container conversion?); unlikely ROI
4. **Compatibility risk** — PowerPoint has limited video codec support (MP4 H.264 primary); re-encoding risky

**Feasibility (Analysis-Only): ⚠️ MEDIUM (Medium Confidence)**

**What Analysis-Only Could Look Like:**

```csharp
public class VideoAnalysis
{
    public string Format { get; set; } // e.g., "mp4"
    public string Codec { get; set; } // e.g., "h264", "vp9"
    public int? Width { get; set; }
    public int? Height { get; set; }
    public int? Bitrate { get; set; } // kbps
    public int? Duration { get; set; } // seconds
    public bool IsOptimizedForPowerPoint { get; set; }
    public string Message { get; set; }
}

// Approach 1: MediaInfo CLI
private VideoAnalysis AnalyzeVideoWithMediaInfo(MediaDataPart videoPart)
{
    // 1. Extract video to temp file
    // 2. Run: mediainfo --Output=JSON video.mp4
    // 3. Parse JSON for codec, dimensions, bitrate
    // 4. Check if codec is PowerPoint-compatible
}

// Approach 2: FFProbe CLI
private VideoAnalysis AnalyzeVideoWithFFProbe(MediaDataPart videoPart)
{
    // 1. Extract video to temp file
    // 2. Run: ffprobe -v quiet -print_format json -show_format -show_streams video.mp4
    // 3. Parse JSON
}

// Approach 3: Pure .NET (Limited)
private VideoAnalysis AnalyzeVideoMinimal(MediaDataPart videoPart)
{
    // Only accessible without external tools:
    // - MIME type (e.g., "video/mp4")
    // - File size
    // - First bytes (may contain container metadata)
    // 
    // Cannot determine: codec, bitrate, duration, dimensions without parsing binary
}
```

**Recommendation for Phase 4:**
- **Skip issue #86 entirely** — rationale: analysis-only is low-value without optimization follow-up
- **If later video optimization is needed:** Add as Phase 5 task; scope would include:
  - MediaInfo/FFProbe CLI integration (external dependency)
  - Codec detection and PowerPoint compatibility check
  - Optional: FFmpeg re-encoding workflow (very heavy)

**Dependencies:**
- **MediaInfo:** CLI tool (external, ~5 MB)
- **FFProbe:** Part of FFmpeg (~50 MB, large)
- Neither recommended for Phase 4; defer to Phase 5

---

## Current pptx-mcp State Analysis

**How pptx-mcp accesses the package:**

1. **PresentationDocument.Open()** is the entry point — opens PPTX via OpenXML SDK
2. **Package access:** `doc.PresentationPart.OpenXmlPackage.Package` gives OPC (Open Packaging Convention) container access
3. **No direct ZipArchive access** — SDK hides it, but can be accessed if needed
4. **Parts enumeration:** `package.GetParts()` returns `PackagePart` objects (logical view, not ZIP entries)

**Current service methods:**
- `GetSlides()` — enumerates SlideParts via PresentationPart
- `GetLayouts()` — enumerates SlideMasterParts and SlideLayoutParts
- `AddSlide()` — creates new SlidePart with layout
- `GetChartData()`, `UpdateChartData()` — chart analysis (chart-specific)

**Pattern observation:**
All current methods use `PresentationDocument.Open()` with `false` (read-only) or `true` (editable) based on mutation needs. This is the correct pattern.

---

## Dependencies Assessment

**Current:** `DocumentFormat.OpenXml` v3.3.0

**Proposed additions:**
- **#80–#84:** Zero new dependencies (standard .NET APIs only)
- **#85 (Image compression):** `SixLabors.ImageSharp` (optional, via configuration)
- **#86 (Video analysis):** MediaInfo or FFProbe CLI (external, not NuGet; skip for Phase 4)

**Recommendation:**
- **Keep Phase 4 dependency-light** — add ImageSharp only if issue #85 is prioritized
- **DocumentFormat.OpenXml 3.3.0 is sufficient** — no need to upgrade to 3.5.0 for Phase 4 work (3.5.0 adds minor schema updates, not required)

---

## Validation & Testing Recommendations

**For each feature:**

1. **File size breakdown (#80):**
   - Test with 10MB+ PPTX (large presentation)
   - Verify ZIP compression ratios match actual disk usage
   - Confirm read-only access (file unchanged)

2. **Media analysis (#81):**
   - Test with mixed media (PNG, JPEG, MP4, audio)
   - Verify hash consistency (same image, different slides)
   - Edge case: images in chart parts

3. **Layout detection (#82):**
   - Test with multi-master presentations
   - Verify accurate usage counts
   - Confirm no false positives (layouts referenced in metadata)

4. **Layout removal (#83):**
   - Test removal → PowerPoint round-trip (must open without warnings)
   - Verify relationship cleanup (no orphaned parts)
   - Atomic failure case (partial deletion recovery)

5. **Deduplication (#84):**
   - Test with 3+ copies of same image
   - Verify relationships redirect correctly
   - PowerPoint round-trip validation
   - Visual verification of images on all slides

6. **Image compression (#85):**
   - Test JPEG/PNG with various quality levels
   - Verify visual quality acceptable
   - Confirm file size reduction ~30–50% for JPEG
   - PowerPoint round-trip validation

---

## Reference File Paths (Prior Art)

**MarpToPptx:**
- `src/MarpToPptx.Pptx/Rendering/OpenXmlPptxRenderer.cs` — PackagePart enumeration, ZipArchive usage, NormalizePackage
- `src/MarpToPptx.Pptx/Extraction/PptxMarkdownExporter.Media.cs` — Image/media enumeration, SHA256 hashing, image dimension detection
- `src/MarpToPptx.Pptx/Diagnostics/TemplateDoctor.cs` — Layout/master validation patterns

**pptx-mcp:**
- `src/PptxMcp/Services/PresentationService.cs` — Current implementation, layout enumeration (lines 96–111)
- `src/PptxMcp/Services/PresentationService.Charts.cs` — Chart data patterns (reference for similar structure)

---

## Conclusion & Recommendations

**Phase 4 Feasibility Summary:**

| Issue | Feature | Feasibility | Confidence | Dependencies | Priority |
|-------|---------|-------------|-----------|---|----------|
| #80 | File size breakdown | ✅ HIGH | HIGH | None | High |
| #81 | Media analysis | ✅ HIGH | HIGH | None | High |
| #82 | Unused layout detection | ✅ HIGH | HIGH | None | High |
| #83 | Layout removal | ⚠️ MEDIUM | MEDIUM | None (test req'd) | Medium |
| #84 | Media deduplication | ⚠️ MEDIUM-HIGH | MEDIUM | None | Medium |
| #85 | Image compression | ⚠️ MEDIUM | MEDIUM | Optional: ImageSharp | Medium-Low |
| #86 | Video analysis | 🤔 MEDIUM | MEDIUM | MediaInfo/FFProbe CLI | Low (defer) |

**Recommended Phase 4 Sequence:**
1. **Start with #80–#82** (pure analysis, no risk, high confidence)
2. **Follow with #83–#84** (mutations, requires PowerPoint testing)
3. **Optional: #85** (if optimization ROI clear; adds ImageSharp dependency)
4. **Defer #86** (low value without optimization; external CLI overhead)

**Next Steps:**
- Cheritto: Pick up implementation for #80–#82 (straightforward)
- Shiherlis: Design test harness (PowerPoint round-trip validation for #83–#84)
- Team: Decide on image compression scope (if #85 included, scope ImageSharp integration early)

---

## Implementation Notes for Implementers

**Pattern to follow:**
- All Phase 4 tools should be in `PresentationService` (or new `PresentationService.Optimization.cs` partial)
- Use existing `GetSlidePart()` / `GetSlideIds()` helpers from service
- Return structured result objects (not strings) — JSON serialization via models
- Follow MCP tool naming: `{Noun}{Verb}` (e.g., `PresentationAnalyzeSize`, `MediaListAssets`)
- Add XML doc comments for MCP SDK Description generation
- All methods read-only unless noted (pass `false` to `PresentationDocument.Open()`)

**Error handling:**
- Wrap `PresentationDocument.Open()` in try/catch — handle corrupted PPTX gracefully
- For mutations (#83, #84): Use backup pattern (copy → modify → verify)
- Return `Success: false` with meaningful Message for all failure cases

---

**Prepared by:** Nate, Consulting Dev  
**Date:** 2026-03-24  
**Status:** Ready for implementation planning


---

### Phase 4: Image Compression Architecture (2026-03-23 McCauley)

**Lead:** McCauley  
**Status:** Active

#### Phase 4: Presentation Optimization

**Lead:** McCauley  
**Status:** Active  
**Created:** 2026-03-19

---

## Executive Summary

Phase 4 ("Presentation Optimization") delivers a comprehensive suite of tools for analyzing and optimizing PowerPoint file sizes. Jon regularly needs to shrink decks by removing unused masters, deduplicating media, and compressing images. Phase 4 is structured as a natural follow-on to Phase 3 (Deck Authoring), focusing on cleanup and optimization of existing presentations.

The phase is scoped into **Tier 1 (read-only analysis)**, **Tier 2 (write operations)**, and **Tier 3 (future/deferred)** to prioritize low-risk, high-value analysis tools first, then enable cleanup operations once the analysis foundation is solid.

---

## Phase 4 Scope & Tier Structure

### Tier 1 — Read-Only Analysis (Low Risk, High Value)

**Principle:** Start with diagnostic tools. These are safe, idempotent, and provide the foundation for write operations.

| # | Title | Description | Complexity | Dependencies |
|---|-------|-------------|-----------|---|
| P4-1 | Analyze presentation file size breakdown | Scan PPTX ZIP structure; report sizes by category (slides, images, video/audio, masters, layouts, other) | Medium | None |
| P4-2 | List and analyze media assets | Enumerate images, video, audio; report size, content type, which slides reference; detect duplicates by content hash | Medium | None |
| P4-3 | Find unused slide masters and layouts | Cross-reference masters/layouts against actual slide usage; report unused and space impact | Medium | None |

**Why Tier 1 First:**
- Safe to ship (no mutations)
- Diagnostic value: users can understand bloat before cleanup
- Enables user-driven cleanup workflows (agent says "remove these masters," user approves)
- Foundation for Tier 2 write operations

---

### Tier 2 — Write Operations (Cleanup & Optimization)

**Principle:** Implement removal and deduplication after analysis is complete. Includes PowerPoint round-trip validation.

| # | Title | Description | Complexity | Dependencies | Risk Mitigation |
|---|-------|-------------|-----------|---|---|
| P4-4 | Remove unused slide masters and layouts | Delete unused masters/layouts from P4-3 analysis; preserve relationship integrity; validate with OpenXmlValidator | Large | P4-3 | OpenXmlValidator before/after; PowerPoint round-trip test |
| P4-5 | Deduplicate identical media | Find media with identical content (SHA256); consolidate to single canonical copy; remove orphans | Large | P4-2 | Relationship validation; round-trip test on duplicated media |
| P4-6 | Compress/optimize images | Downscale images larger than display; target DPI selection; format conversion (BMP/TIFF→PNG/JPEG); compression stats | Large | None | SkiaSharp dependency decision; JPEG quality tuning; round-trip validation |

**Why Tier 2 After Tier 1:**
- Tier 1 analysis runs first; users see what can be cleaned
- Tier 2 operations are higher-risk (package mutation); require validated analysis
- Each tool preserves package integrity and tests PowerPoint compatibility

---

### Tier 3 — Future/Deferred

| # | Title | Description | Scope | Deferred Reason |
|---|-------|-------------|-------|---|
| P4-7 | Optimize embedded video (Analysis Only) | Enumerate videos; report codec, resolution, bitrate, duration; suggest compression (no transformation) | Analysis only | Video re-encoding requires ffmpeg/external tool; deferred to future spike |

---

## GitHub Issues Created

All 7 issues created under **Phase 4: Presentation Optimization** milestone (GitHub milestone #5), labeled with `squad` and `phase-4`.

### Issue Mapping

| GitHub # | Title | Tier | Label | Status |
|---|---|---|---|---|
| #80 | P4-1: Analyze presentation file size breakdown | Tier 1 | analysis, type:feature | Ready |
| #81 | P4-2: List and analyze media assets | Tier 1 | analysis, media, type:feature | Ready |
| #82 | P4-3: Find unused slide masters and layouts | Tier 1 | analysis, type:feature | Ready |
| #83 | P4-4: Remove unused slide masters and layouts | Tier 2 | optimization, type:feature | Blocked on P4-3 |
| #84 | P4-5: Deduplicate identical media | Tier 2 | optimization, media, type:feature | Blocked on P4-2 |
| #85 | P4-6: Compress/optimize images | Tier 2 | optimization, media, type:feature | Ready |
| #86 | P4-7: Optimize embedded video | Tier 3 | analysis, media, type:feature, go:no | Deferred |

---

## Key Architectural Decisions

### 1. Read-Only Analysis First

**Decision:** Implement all Tier 1 analysis tools before any Tier 2 write operations.

**Rationale:**
- Analysis tools are safe (no mutations); establish confidence in package scanning
- Provide diagnostic value immediately
- Foundation for write operations (Tier 2 relies on Tier 1 patterns)
- Enables user-in-the-loop workflows (agent shows analysis, user approves cleanup)

---

### 2. Image Compression: SkiaSharp Dependency

**Decision:** Use **SkiaSharp** for image downscaling and format conversion (P4-6).

**Rationale:**
- **Cross-platform:** Works on Windows, macOS, Linux
- **High quality:** Proven library, used in production
- **Modern:** Active development, good .NET integration
- **Alternative considered:** System.Drawing.Common (Windows-only in .NET 6+, legacy)

**Implementation Notes:**
- Add NuGet dependency: `SkiaSharp` (latest stable, currently ~2.88)
- Update README to document image optimization capabilities and SkiaSharp requirement
- Document JPEG compression quality setting: recommend 85% as starting point (tunable via parameter)
- Add SkiaSharp to CI dependencies if needed

---

### 3. OpenXML Validation & PowerPoint Round-Trip

**Decision:** All Tier 2 write operations must:
1. Run `OpenXmlValidator` before and after
2. Include PowerPoint round-trip test (save modified PPTX, open in PowerPoint, verify fidelity)

**Rationale:**
- A file can pass `OpenXmlValidator` and still fail to open in PowerPoint (learned from Phase 1/2 experience)
- PowerPoint compatibility is the real success criterion
- Round-trip tests catch subtle package structure issues

**Implementation Pattern:**
```csharp
// Before operation
var validator = new OpenXmlValidator();
var beforeErrors = validator.Validate(presentationPart);

// Perform operation (e.g., remove master)
// ...

// After operation
var afterErrors = validator.Validate(presentationPart);

// Return results with validation status
return new RemoveResult 
{ 
    ItemsRemoved = removed.Count,
    SpaceSaved = savedBytes,
    ValidationBefore = new { ErrorCount = beforeErrors.Count },
    ValidationAfter = new { ErrorCount = afterErrors.Count },
    Success = afterErrors.Count == 0
};
```

---

### 4. Media Hash-Based Deduplication

**Decision:** Use SHA256 content hash for media deduplication (P4-5).

**Rationale:**
- Deterministic: same content always produces same hash
- Reliable: SHA256 collisions extremely unlikely
- Simple to implement: read stream, compute hash, compare
- Proven pattern: used in backup/dedup systems

**Implementation:**
```csharp
using var sha256 = System.Security.Cryptography.SHA256.Create();
var hash = sha256.ComputeHash(mediaStream);
var hexHash = Convert.ToHexString(hash);
```

---

### 5. ZIP-Level Package Scanning (P4-1)

**Decision:** Use `System.IO.Compression.ZipArchive` to enumerate package contents for size analysis.

**Rationale:**
- PPTX is a ZIP; direct access gives complete visibility
- `ZipArchiveEntry.Length` provides exact sizes
- Complementary to OpenXML SDK (which understands semantics)

**Implementation Pattern:**
```csharp
using (var archive = ZipFile.OpenRead(filePath))
{
    foreach (var entry in archive.Entries)
    {
        var category = CategorizeByContentType(entry.Name);
        totalByCategory[category] += entry.Length;
    }
}
```

---

## Recommended Implementation Order

### Rationale

1. **P4-1, P4-2, P4-3 (Tier 1 analysis):** Implement in any order; they are independent. Each unblocks corresponding Tier 2 tool.
   - **Suggested sequence:** P4-1 → P4-2 → P4-3 (building complexity)
   
2. **P4-4 (depends on P4-3):** Remove masters/layouts. Implement after P4-3 analysis is complete and tested.

3. **P4-5 (depends on P4-2):** Deduplicate media. Implement after P4-2 media enumeration is solid.

4. **P4-6 (independent):** Image compression. Can be implemented in parallel with Tier 2; includes new SkiaSharp dependency.

5. **P4-7 (Tier 3, deferred):** Video analysis. Mark as low priority; consider for post-Phase-4 spike if demand exists.

---

## Team Assignment & Capacity

**Default Assignment (provisional, subject to squad prioritization):**

| Phase | Tool | Owner | Effort | Notes |
|-------|------|-------|--------|-------|
| Tier 1 | P4-1 (File size analysis) | Cheritto | 3–4h | ZIP enumeration + OpenXML categorization |
| Tier 1 | P4-2 (Media analysis) | Cheritto | 3–4h | Hashing, reference tracking, dedup detection |
| Tier 1 | P4-3 (Unused masters) | Cheritto | 3–4h | Master/layout traversal + usage cross-ref |
| Tier 2 | P4-4 (Remove masters) | Cheritto | 5–6h | Relationship updates, OpenXmlValidator, round-trip |
| Tier 2 | P4-5 (Dedup media) | Cheritto | 5–6h | Relationship updates, orphan cleanup, round-trip |
| Tier 2 | P4-6 (Compress images) | Cheritto | 6–8h | SkiaSharp integration, DPI logic, JPEG tuning, round-trip |
| **Total Tier 1+2** | | | **~32 hours** | 2–3 weeks part-time for one dev |
| Tier 3 | P4-7 (Video analysis) | TBD | 2–3h | Deferred; consider for future spike |

**Testing & Documentation (parallel):**
- **Shiherlis:** E2E tests for Tier 2 operations (P4-4, P4-5, P4-6); round-trip validation on real presentations
- **@copilot:** Document Phase 4 tools, examples, video optimization analysis patterns
- **Nate (optional):** Code review for OpenXML package integrity and SkiaSharp integration

---

## Success Criteria

✅ **All Tier 1 issues closed:** Analysis tools available, tested, documented  
✅ **All Tier 2 issues closed:** Write operations complete, PowerPoint compatibility verified  
✅ **OpenXmlValidator passes:** Before/after validation on all write operations  
✅ **E2E round-trip tests pass:** Modified PPTX files open correctly in PowerPoint  
✅ **Test coverage:** 3+ unit test cases per tool, comprehensive edge cases  
✅ **Documentation:** README updated, TOOL_REFERENCE.md and EXAMPLES.md include Phase 4 tools  
✅ **No Phase 3 regression:** All existing tests continue to pass  

---

## Known Risks & Mitigation

| Risk | Impact | Mitigation |
|------|--------|-----------|
| SkiaSharp dependency adds complexity | Build/packaging | Document in README; add to CI; plan for NuGet publishing |
| Image downscaling affects visual quality | User experience | Test JPEG quality settings (start at 85%); log before/after metadata |
| Master/layout removal could orphan slides | Data loss | Validate relationships before removal; test round-trip with real presentations |
| Media dedup removes wrong copy | Data loss | Use content hash (SHA256); test with intentionally duplicated media |
| OpenXmlValidator passes but PowerPoint fails | Hidden bugs | Require round-trip tests for all Tier 2 operations |

---

## Future Enhancements (Post-Phase 4)

1. **Video re-encoding (P4-7 full scope):** Requires ffmpeg or similar; spike with architecture decision
2. **Batch optimization:** Multi-operation pipeline (analyze → deduplicate → compress → remove)
3. **Optimization presets:** "Aggressive" (downscale to 150 DPI, remove all unused), "Conservative" (remove masters only)
4. **Optimization report:** JSON summary of optimization opportunities and actual space saved
5. **MCP UX improvements:** Async tasks for large presentations, progress notifications

---

## Decision Log

- **2026-03-19:** Phase 4 scope defined, tier structure approved, 7 issues created
- **2026-03-19:** GitHub milestone #5 created; labels added
- **2026-03-19:** SkiaSharp chosen for image optimization (vs. System.Drawing.Common, ImageSharp)
- **2026-03-19:** OpenXML validation + round-trip testing established as acceptance criteria for Tier 2

---

## Recent Decisions (March 26, 2026)

### Tool Consolidation Implementation (PR #97, Issue #92)

**Lead:** Cheritto  
**Date:** 2026-03-26  
**Status:** ✅ Implemented (PR #97 merged)

Implemented Tier 1 consolidation + `pptx_update_text` deprecation from McCauley's analysis (#92):

**Consolidations:**
1. `pptx_find_unused_layouts` + `pptx_remove_unused_layouts` → `pptx_manage_layouts` (action enum: Find | Remove)
2. `pptx_analyze_media` + `pptx_deduplicate_media` → `pptx_manage_media` (action enum: Analyze | Deduplicate)
3. `pptx_update_text` deprecated — `pptx_update_slide_data` is a strict superset

**Rationale:**
- Follows McCauley's analysis in docs/TOOL_CONSOLIDATION_ANALYSIS.md
- Domain-specific pairings (analyze → act) are more natural for LLM tool selection
- Clean break, no deprecation period (old tool names never shipped as stable API)
- Service layer unchanged, minimizing risk

**Impact:**
- Tool count: 24 → 21
- Service-layer tests: all pass unchanged
- Build: 0 errors
- Tests: 552 green

**Files:**
- Created: ManageLayoutsAction.cs, ManageMediaAction.cs, PptxTools.ManageMedia.cs
- Modified: PptxTools.cs, PptxTools.Optimization.cs, README.md, tests
- Deleted: PptxTools.Media.cs, PptxTools.Deduplication.cs

---

### Media Deduplication Implementation Pattern (Issue #84)

**Lead:** Cheritto  
**Date:** 2026-03-24  
**Status:** ✅ Implemented

**Decision:** Use `ownerPart.DeletePart(duplicatePart)` for ImagePart deduplication, not `DeleteReferenceRelationship`.

**Context:** When redirecting ImagePart references from a duplicate to a canonical part, we need to remove the old relationship from each owner part.

**Rationale:**
- `DeletePart()` correctly handles standard OpenXmlPart relationships (ImagePart)
- `DeleteReferenceRelationship()` is only for DataPart references (video/audio)
- Using the wrong method throws at runtime

**Impact:** Future media deduplication work (video/audio DataParts) will use `DeleteReferenceRelationship` for those part types.

---

### Write Operation Validation Pattern for Layout Removal (Issue #83)

**Lead:** Cheritto  
**Date:** 2026-03-24  
**Status:** ✅ Established

All write-operation optimization tools should follow this pattern:

1. **Read-only analysis first** — reuse the corresponding analysis tool (e.g., FindUnusedLayouts) to identify targets before opening writable
2. **OpenXmlValidator before AND after** — capture error counts at both checkpoints and surface them via `ValidationStatus`
3. **Safety intersection** — when caller provides explicit targets, intersect with the actually-unused set rather than trusting blindly
4. **Cleanup ID lists** — when deleting parts (layouts, masters), always remove corresponding ID entries from parent XML lists (SlideLayoutIdList, SlideMasterIdList) before calling DeletePart

**Rationale:** PowerPoint is the real validator. OpenXmlValidator catches structural issues early but doesn't guarantee PowerPoint compatibility. Two-phase approach (analyze read-only, then modify) minimizes writable lock time and keeps safety check separate from mutation.

---

### User Directive: Image Optimization Library (Issue #85)

**Author:** Jon Galloway (via Copilot)  
**Date:** 2026-03-24  
**Status:** ✅ Approved

For issue #85 (compress/optimize images), use **Magick.NET** instead of ImageSharp.

**Rationale:** User request — captured for team memory.

---

### Magick.NET Research & Feasibility (Issue #85)

**Lead:** Nate (Consulting Dev)  
**Date:** 2026-03-26  
**Status:** ✅ Completed — GO verdict

**Verdict: GO** — Magick.NET is **fully viable** for issue #85.

**Capabilities (all supported):**
- Read image dimensions
- Downscale images maintaining aspect ratio
- Convert BMP/TIFF → PNG/JPEG
- Re-encode JPEG at configurable quality
- Stream-based I/O (no temp files)

**Recommendation:**
- Use `Magick.NET-Q8-x64` (version 14.11.0+)
- Q8 = 8 bits per pixel (sufficient for PPTX images)
- x64 = platform-specific (reduces binary size ~15-18 MB)

**Bundle Size Impact:** +15-35 MB added to published binary (acceptable for open-source MCP server)

**Cross-Platform:** Full support on Windows, Linux (ubuntu-latest), macOS via native binary bundling

**Integration Pattern:** Create `PresentationService.ImageOptimization.cs` with:
- `OptimizeImages()` public tool method
- `OptimizeImagesInSlidePart()` helper
- `OptimizeImagePart()` Magick.NET wrapper

**Timeline:** 6–8 hours (dependency setup, tool implementation, E2E test, documentation)

**vs. SkiaSharp:** Magick.NET is better for image compression (format conversion, JPEG quality control) despite being slower. SkiaSharp better for real-time rendering.

**Key Gotchas:**
- Preserve aspect ratio by setting unused dimension to 0 in `Resize()`
- Use `imagePart.FeedData(stream)` to replace image in-place
- Add `.csproj` properties to ensure native binaries copied during publish on Linux

**Next Step:** Pass research to Cheritto with implementation sketch for development.


---

## Decision Archive: Merged from Inbox (2026-03-25)

### CLI Interface Decomposition (#94 → #98–#105)

**Lead:** McCauley  
**Date:** 2026-03-24  
**Status:** ✅ Approved — 7 GitHub issues created  

The design decomposes #94 into 7 implementable sub-issues. Dual-mode architecture approved: pptx-tools --stdio (MCP server) vs pptx-tools analyze [options] (CLI). System.CommandLine 3.0 for CLI surface. 21 MCP tools reused without refactoring.

**Key decisions:**
- Single binary, dual-mode entry point (40 lines of Program.cs)
- 7 command groups: analyze, optimize, inspect, export, edit, media, slides
- Compound command (optimize) orchestrates multiple service calls with reporting
- NuGet Phase 1, Scoop/Homebrew Phase 2

**Effort:** 22–32 hours over 3 weeks. All 7 issues await #98 (foundation). See full decision document in .squad/decisions/inbox/mccauley-cli-decomposition.md (archived).

---

### System.CommandLine v3 CLI Command Pattern (Analysis)

**Author:** Cheritto  
**Date:** 2026-03-24  
**PR:** #108 (squad/99-102-analyze-export)

Implementing CLI commands on System.CommandLine v3.0.0-preview.2. The v3 API differs from v2 docs online.

**Key patterns:**
- Command factory: static class with Create(PresentationService) returning fully-configured Command
- v3 API: Argument<T>("name") constructor, set Description via property; Option<T>("--name") likewise
- SetAction receives ParseResult, NOT context object
- Exit codes: cast to Func<ParseResult, int> for return-value overload
- Value access: parseResult.GetValue(argObj) / parseResult.GetValue(optObj)
- DI: lightweight ServiceCollection → BuildServiceProvider in RunCliAsync

**Impact:** All future CLI command implementations (#99–#105) follow this pattern. Reference for team.

---

### System.CommandLine v3 CLI Command Pattern (Inspect/Media)

**Author:** Cheritto  
**Date:** 2026-03-27  
**Context:** PR #109 - CLI inspect, media, slides commands

**Established pattern:**
- Constructors: single-arg with Description property
- SetAction takes ParseResult parameter, use parseResult.GetValue(argObj)
- Exit codes: Use Environment.ExitCode (not ctx.ExitCode)
- DI: ServiceCollection + PresentationService singleton in RunCliAsync
- JSON output: --json flag on read-only subcommands
- Slide numbering: CLI takes 1-based --slide, converts to 0-based for GetSlideContent/GetSlideXml

**Consistency:** Matches AnalyzeCommand/ExportCommand patterns established in #108.

---

### Dual-Mode Entry Point Implementation

**Author:** Cheritto  
**Date:** 2026-03-27  
**Status:** ✅ Implemented (PR #107)

Refactored Program.cs to support both MCP server mode (--stdio) and CLI mode from the same binary.

**Key decisions:**
- System.CommandLine 3.0.0-preview.2 (latest resolved by dotnet add --prerelease; v3 API differs from v2)
- Mode detection: string-based DetermineMode() returns "mcp" or "cli" (simple, testable)
- MCP path byte-for-byte identical (zero risk to existing functionality)
- Stub commands use SetAction with Console.WriteLine (minimal surface for #99–#105 to replace)

**Impact:** Foundation for all CLI issues (#99–#105). Zero MCP regressions (575/575 tests pass). 2 files changed, +81/-21 lines.

---

### CLI edit command — JSON input patterns

**Author:** Cheritto  
**Date:** 2026-03-27  
**Context:** Issue #103, PR #112

**Two JSON input patterns:**
1. **File-based** (edit batch): mutations loaded from JSON file path (large payloads, composable)
2. **Inline** (table, chart): JSON passed directly as option values (small inputs, single-command convenience)

All JSON deserialization uses PropertyNameCaseInsensitive = true.

**Team pattern:** File-based for unbounded input, inline for bounded/small structured input.

---

### OptimizeCommand compound CLI pattern

**Author:** Cheritto  
**Date:** 2026-03-24  
**Issue:** #100  
**PR:** #111  

The optimize command chains multiple service calls (dedup, image compress, layout removal) in sequence.

**Decisions:**
1. Copy-first: File.Copy before any mutations; original file untouched
2. Try/catch per step: each optimization step independently wrapped; if dedup fails, others still run
3. --no-* toggle pattern: --remove-layouts defaults true; --no-remove-layouts explicitly disables (System.CommandLine v3 lacks native negation support)
4. FormatBytes duplicated: copy from AnalyzeCommand; extract to CliHelpers if a third command needs it
5. JSON output model private to command: OptimizeStepResult and OptimizeResult nested in OptimizeCommand class

**Team impact:** The edit command (#103) is last remaining stub. --no-* toggle pattern reusable for future boolean defaults-to-true options.

---

### Video/Audio Metadata Extraction Research (Issue #86)

**Consulting Dev:** Nate  
**Date:** 2026-03-24  
**Deliverable:** NuGet package recommendation for embedded video/audio metadata extraction

**Verdict: GO with SharpMp4Parser** ⭐

After evaluating 10+ candidates, **SharpMp4Parser v0.0.5** is the only solution meeting all requirements: MIT license, pure .NET, cross-platform (Windows/Linux/macOS), zero native dependencies, stream-based API, extracts codec/resolution/bitrate/duration, NuGet-native (~100KB).

**Rejected alternatives:**
- FFMpegCore, MediaInfo.DotNetWrapper, FFMediaToolkit: ALL require external native binaries (FFmpeg, MediaInfo) — incompatible with MCP server distribution
- TagLibSharp: Pure .NET but inadequate for VIDEO (excellent for audio; no video codec/resolution extraction)

**Implementation pattern:** Analogous to Magick.NET — lightweight .NET wrapper focused on one job (metadata extraction), works from streams, zero external dependencies, MIT-licensed, cross-platform by design.

**Complexity:** 200–300 lines of helper methods to map MP4 box structure → codec/resolution/bitrate/duration. Prototype within Services/PresentationService.cs.

**Effort:** 7–10 hours (Phase 1: dependency + box parsing helpers; Phase 2: MCP tool; Phase 3: tests).

**Escalation:** If 10% of presentations fail (malformed MP4), escalate to FFMpegCore + binary distribution discussion.

---

### CLI Mode Detection Design (Copilot Directive)

**Date:** 2026-03-24T20:17:00Z  
**Decision Authority:** Jon Galloway (via Copilot)

--stdio flag for MCP server mode. No args = show help (standard CLI convention). No backward compatibility concern — pptx-tools is not yet released. The serve subcommand approach is dropped.

**Rationale:** --stdio matches MCP ecosystem convention, no-args-help matches universal CLI convention (dotnet, git, gh), and since the tool isn't released yet, backward compat is irrelevant.

---

### Project Rename Decision (Copilot Directive)

**Date:** 2026-03-24T21:50:00Z  
**Decision Authority:** Jon Galloway (via Copilot)

Rename project from pptx-mcp to pptx-tools (repo + NuGet package). CLI command becomes pptx. Namespace changes from PptxMcp to PptxTools. Execute after CLI commands are complete.

**Rationale:** Broader discoverability for non-MCP users searching for PowerPoint CLI tools. MCP is a feature, not the identity.

---

## End Inbox Archive (2026-03-25T04:50:46Z)

All 10 inbox files merged above and deleted from .squad/decisions/inbox/.


---

### Issue #114 Orchestration: Hyperlink Support Implementation

**Orchestrator:** McCauley  
**Date:** 2026-03-25  
**Status:** ✅ Implemented (PR #145)

**Parallel Agents:** Shiherlis (testing), Cheritto (feature)

**Work Summary:**
- **Shiherlis:** 49 comprehensive tests (30 service + 19 tool) covering all hyperlink CRUD operations
- **Cheritto:** Full feature implementation — HyperlinkInfo model, PresentationService.Hyperlinks.cs service layer, pxtx_manage_hyperlinks tool (Get/Add/Update/Remove actions)
- **Build:** 0 errors, **Tests:** 624/624 passing
- **New Files:** 5 (HyperlinkInfo.cs, PresentationService.Hyperlinks.cs, PptxTools.Hyperlinks.cs, HyperlinkTests.cs, HyperlinkToolsTests.cs)

**Impact:** Feature-complete hyperlink support. Ready for review & merge.

