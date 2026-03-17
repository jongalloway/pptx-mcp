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
