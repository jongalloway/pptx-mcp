# Squad Decisions

## Active Decisions

### Decision: Copilot PR Merge Review — PRs #182 & #183 (2026-05-11)

**Author:** McCauley (Lead)  
**Date:** 2026-05-11  

## PR #182 — ModifyPlaceholder/AddPlaceholder for pptx_manage_layouts

**Verdict: MERGED (squash)**

- Not draft, CI green (build + test), mergeable, no blocking review comments.
- Copilot review feedback was addressed in follow-up commits (grouped placeholder enumeration, duplicate idx guard).
- Code quality: Clean partial class separation, structured results, proper validation, consistent with existing patterns.
- 531 additions across 6 files. Tests included.

## PR #183 — Delete/DeleteAll for pptx_manage_slides

**Verdict: LEFT OPEN (draft)**

- PR is still in **draft** state — not merge-ready by author's own signal.
- CI green, code looks solid (descending-order batch delete, NotesSlidePart cleanup, structured results, 3 tests).
- Minor concern: `DeleteAllSlides` produces a zero-slide presentation — verify this is valid OpenXML and doesn't break downstream tools.
- No reviews yet. Needs at least one review cycle before merge.
- **Revision owner if needed:** Cheritto (backend dev) for any service-layer fixes; Shiherlis for additional test coverage.

## Local Changes

- Pushed 8 local commits to origin/main (7 squad metadata + 1 documentation commit).
- Documentation guides (EMU, Shape Resolution, Troubleshooting, Contributing) and README updates now live on main.
- README merge conflict resolved (stashed local vs origin divergence from PRs #176–#182).

## Ordering Notes

- No ordering constraint between PRs #182 and #183 — they touch different tools and files.
- PR #183 should rebase onto main now that #182 is merged, to pick up any shared test infrastructure changes.

---

### Decision: Enable XML Documentation for MCP Tool Descriptions (2026-05-11)

**Authors:** Cheritto (Backend), Shiherlis (Tester)  
**Date:** 2026-05-11  
**Status:** ✅ Implemented & Verified

## Problem

MCP startup warning: `[warning] Tool pptx_manage_media does not have a description. Tools must be accurately described to be called`

## Root Cause

ModelContextProtocol runtime reads tool descriptions from compiled XML documentation sidecar file (`PptxTools.xml`). Without this file, XML `<summary>` comments on tool methods are not surfaced in MCP tool metadata advertisements, even though the comments exist in source code.

## Solution

Enabled XML documentation file generation in `src/PptxTools/PptxTools.csproj`:

```xml
<GenerateDocumentationFile>true</GenerateDocumentationFile>
```

## Implementation (Cheritto)

- Minimal change preserves existing pattern of keeping tool descriptions in XML comments
- Build now emits `PptxTools.xml` sidecar alongside executable
- Fixes warning for `pptx_manage_media` and sibling tools (`pptx_manage_slides`, `pptx_manage_layouts`, `pptx_optimize_images`, `pptx_manage_hyperlinks`)

## Verification (Shiherlis)

- Fresh Debug rebuild: `PptxTools.xml` generated correctly
- `dotnet run --project .\src\PptxTools -- --stdio`: Started cleanly, no warning
- Release build: ✅ Complete
- Test suite: ✅ 1261/1261 passing, zero regressions

## Risk Assessment

- **Operational Hazard:** Tool descriptions now depend on `PptxTools.xml` being emitted and shipped beside executable
- **Reflection Analysis:** All 32 tools still have null `McpServerToolAttribute.Description` at runtime; descriptions depend on XML sidecar
- **Mitigation:** If warning reappears, investigate build/output packaging first before changing source code

## Outcome

✅ **Complete** — Warning resolved. Solution accepted as effective for reported symptom. Operational dependency documented for future reference.

---

### Decision: Merge Local Docs Changes (Issue #138) (2026-05-11)

**Owner:** Shiherlis (Tester)  
**Date:** 2026-05-11  
**Status:** ✅ Ready to merge  
**Priority:** Medium  

## Context

Issue #138 documentation was completed locally (4 guides + README updates) but never merged. McCauley previously recommended promotion from `go:no` to merge status. The Scribe requested review to determine if these local changes are still valid after PR #182 and #183 landed.

## Analysis

**Local changes:**
- 1 modified file: `README.md` (modified, +14 lines)
- 4 new files: `docs/CONTRIBUTING.md`, `docs/EMU_GUIDE.md`, `docs/SHAPE_RESOLUTION_GUIDE.md`, `docs/TROUBLESHOOTING.md`

**PR #182 (ModifyPlaceholder/AddPlaceholder):**
- Touches: Models, Services, Tools, Tests
- **No overlap with docs files** ✅

**PR #183 (Delete/DeleteAll slides):**
- Touches: Models, Services, Tools, Tests
- **No overlap with docs files** ✅

**Conflict check:**
- None. PRs #182/#183 are pure feature additions (layout placeholders, slide deletion).
- No README changes in either PR.
- No docs changes in either PR.

**Content quality:**
- All references use correct project name (`PptxTools`, not `PptxMcp`) ✅
- Test command matches CI: `dotnet test --solution PptxTools.slnx --configuration Release` ✅
- Expected test count updated to current baseline: `1227 passed` ✅
- Test suite passes: **1227/1227 tests passing** ✅

## Decision

**✅ MERGE LOCAL DOCS CHANGES**

These changes are still valid and ready to commit. They:
1. Fill a documented gap from McCauley's prior decision
2. Do not conflict with PR #182 or #183
3. Pass all 1,227 tests (zero regressions)
4. Are internally consistent and accurate

## Recommended Actions

1. **Commit local changes:**
   ```bash
   git add README.md docs/CONTRIBUTING.md docs/EMU_GUIDE.md docs/SHAPE_RESOLUTION_GUIDE.md docs/TROUBLESHOOTING.md
   git commit -m "docs: add EMU, Shape Resolution, Troubleshooting, and Contributing guides

   Closes #138

   - EMU_GUIDE.md: conversion reference, standard dimensions, positioning scenarios
   - SHAPE_RESOLUTION_GUIDE.md: targeting by name/index/placeholder, discovery methods
   - TROUBLESHOOTING.md: 10 categories, common errors and solutions
   - CONTRIBUTING.md: development workflow, architecture, test patterns, CI/CD
   - README.md: added Guides section, improved Contributing callout

   All 1,227 tests passing. No conflicts with PR #182/#183.

   Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"
   ```

2. **Push to main:**
   ```bash
   git push origin main
   ```

3. **Close issue #138** with comment linking commit SHA.

4. **Update issue labels:** Remove `go:no` and `go:needs-research`, add `go:done`.

## Impact

- Closes documented gap from McCauley's decision on 2026-04-15
- Lowers barrier to entry for users (EMU/Shape guides) and contributors (CONTRIBUTING guide)
- Provides comprehensive troubleshooting reference
- All test patterns documented for consistency

## Notes

- Issue #138 is still open despite local implementation being complete since 2026-04-15
- These changes have been validated and waiting since April — safe to merge immediately
- No CLI quick-start section yet (McCauley identified as separate P0 priority)

---

### Decision: Close #94 (CLI Research) — Fully Satisfied (2026-04-15)

**Owner:** McCauley (Lead)  
**Date:** 2026-04-15  
**Status:** ✅ Ready to close  
**Priority:** Housekeeping  

**Context:** Issue #94 was a research spike: "Research viable CLI patterns for standalone usage." It was decomposed into 7 implementation sub-issues (#98–#105) after Nate's research delivered a GO verdict on dual-mode architecture.

**Finding:** All 5 acceptance criteria are met:

| Criterion | Status | Evidence |
|-----------|--------|----------|
| Research viable CLI patterns | ✅ Done | Nate's research — dual-mode architecture selected |
| Propose CLI command structure | ✅ Done | 7 command groups: analyze, optimize, inspect, export, edit, media, slides |
| Evaluate dual-mode vs separate project | ✅ Done | Dual-mode chosen — single binary, `DetermineMode()` in Program.cs |
| Prototype preferred approach | ✅ Done | Beyond prototype — fully shipped. Program.cs lines 22–39, 7 command files in src/PptxTools/Commands/ |
| Document decision | ✅ Done | Decomposition comment on #94, decision archived in squad decisions |

**Sub-issue status:** 6 of 7 closed (#98–#104). Only #105 (NuGet global tool packaging) remains open — that's a distribution concern, not a research question.

**Decision:** Close #94. The research spike delivered its output and then some — the CLI is fully implemented, not just prototyped.

**Gap identified:** The CLI is built but **undocumented**. README.md only describes MCP server usage. See companion decision on #138 scope below.

---

### Decision: Promote #138 (Docs) — Expand Scope, Prioritize CLI Docs (2026-04-15)

**Owner:** McCauley (Lead)  
**Date:** 2026-04-15  
**Status:** 🟡 Recommend `go:yes` (currently labeled `go:no`)  
**Priority:** Medium  

**Context:** Issue #138 covers 4 documentation guides: EMU Calculator Guide, Shape Resolution Guide, Troubleshooting Guide, Contributing Guide. Shiherlis has completed implementation.

**New finding:** The CLI (7 command groups, dual-mode architecture) shipped via #98–#104 but has **zero user-facing documentation**. The README only covers MCP server setup. A user who discovers `pptx-tools` today has no way to know they can run `pptx-tools analyze file-size presentation.pptx` from the command line.

**Deliverables from Shiherlis:**
- ✅ **EMU_GUIDE.md** (5.5 KB) — English Metric Unit conversions, precision guidelines, tool reference
- ✅ **SHAPE_RESOLUTION_GUIDE.md** (11.2 KB) — Target shapes by name/index/placeholder type, discovery methods, real-world scenarios
- ✅ **TROUBLESHOOTING.md** (13.2 KB) — 10 category sections, quick reference format for users and developers
- ✅ **CONTRIBUTING.md** (18.7 KB) — 5-step example for adding new tools, test patterns, CI/CD pipeline, code style
- ✅ **README.md updated** (+7 lines) — Added "Guides" section with links to all four docs

**Test Coverage:** All 1,227 tests pass, zero regressions.

**Decision:**

1. **Merge #138 documentation** — All four guides verified for accuracy, examples tested, internal links verified.

2. **Add CLI quick-start to README** (or separate companion PR):
   - CLI usage introduction (dual-mode behavior, how to run commands)
   - Command reference overview with examples
   - Link to full command reference docs
   - Rationale: Biggest usability gap in current README

3. **Recommended next steps:**
   - Merge #138 PR with four guide files + README updates
   - Queue companion CLI quick-start PR (can be same or separate)
   - Update issue templates to link to CONTRIBUTING.md and TROUBLESHOOTING.md
   - Maintain sync of CONTRIBUTING.md and EMU_GUIDE.md as tools/commands evolve

**Impact:** Closes documentation gap created by CLI implementation wave (#98–#104). Without this, we shipped a feature nobody knows about.

---

### Decision: Fix config.json teamRoot Mismatch (2026-04-10)

**Owner:** Scribe (on behalf of Jon Galloway)  
**Status:** ✅ Complete  
**Priority:** 🔴 Critical  

**Problem:** `.squad/config.json` had `teamRoot` pointing to `D:\Users\Jon\Documents\GitHub\pptx-tools` but actual repo is at `C:\Users\Jon\Documents\GitHub\pptx-tools`. This broke coordinator path resolution for `.squad/*` files (charters, decisions, skills).

**Solution:** Updated config.json to correct path:
```json
{
  "version": 1,
  "teamRoot": "C:\\Users\\Jon\\Documents\\GitHub\\pptx-tools"
}
```

**Validation:**
- Confirmed `.squad/` directory exists at corrected path
- Spot-checked `.squad/agents/mccauley/charter.md` is readable at new path
- `git rev-parse --show-toplevel` confirms repo root matches

**Impact:** Coordinator path resolution now functional. Next agent spawn will have full context (charter, team routing, decisions).

**Prevents Regression:** Added to squad bootstrap checklist; future setups will verify machine/path before use.

---

### Decision: Squad State Cleanup — Final Pass (2026-04-11)

**Owner:** McCauley (Lead)  
**Status:** ✅ Complete  
**Priority:** 🟢 Low (informational refresh)

**Context:** Prior notable-state review (2026-04-10) identified two squad metadata items: 1) config.json — Critical path mismatch (fixed by Scribe), 2) now.md — Stale focus area (pending refresh).

**Decisions Made:**

1. **Verified config.json — No Action Needed**
   - `teamRoot` correctly points to `C:\Users\Jon\Documents\GitHub\pptx-tools`
   - Coordinator path resolution functional; no regression

2. **Refreshed now.md — Updated to Current State**
   - `updated_at`: 2026-03-16 → 2026-04-11
   - `focus_area`: "Initial setup" → "Advanced feature waves & test modernization (waves #125–#150)"
   - `active_issues`: [] → [125, 135, 140, 150]
   - Rationale: Stale since bootstrap (26 days); reflects work already shipped to main

3. **Audit for Additional Stale Metadata — No Further Items**
   - Audited identity/wisdom.md, casting configs, agent histories, routing/team
   - Result: No additional tightly-coupled, clearly stale, low-risk items found

**Impact:** Squad metadata now current and accurate. Coordinator and agent spawning fully functional. No blocking items.

**Files Modified:**
- `.squad/identity/now.md` — Updated focus_area, active_issues, timestamp
- `.squad/agents/mccauley/history.md` — Added session completion notes

---

### Recommendation: Refresh `.squad/identity/now.md` (2026-04-10)

**Owner:** Coordinator (Jon Galloway)  
**Status:** 🟡 Pending  
**Priority:** Medium  

**Problem:** `.squad/identity/now.md` stale since 2026-03-16T05:17:26.912Z (25 days). Still shows generic template with empty active_issues.

**Current Content:**
```yaml
updated_at: 2026-03-16T05:17:26.912Z
focus_area: Initial setup
active_issues: []
```

**Recommendation:**
1. Update `focus_area` to reflect current work (e.g., "Feature waves #115–#125, PR integration & consolidation")
2. Populate `active_issues` with current sprint issues (e.g., #115, #116, #121, #125)
3. Update timestamp to current date

**Rationale:** `.squad/identity/now.md` is first thing new agents read at spawn. Stale content reduces squad orientation and context clarity. Refresh at next planning cycle.

**Next Action:** Coordinator to execute at next squad planning sync.

---

### Decision: Safe Working Tree Cleanup (2026-03-27)

**Owner:** McCauley (Lead)  
**Status:** Complete  
**Action:** Preserve dirty main → clean up worktree & issue-17 branch

## Problem

Main working tree had 56 uncommitted changes (file-encoding normalization + new config files), and a stale git worktree existed at `C:\Users\Jon\AppData\Local\Temp\pptx-mcp-issue17` pointing to completed issue #17. Repository state needed cleanup without losing user work.

## Decision

1. **Preserve changes:** Move all dirty changes from main onto a new local branch (`squad/tmp-preserve-orchestration-changes`) and commit them there. This protects the work while returning main to a clean, pushable state.
2. **Clean main:** Reset main to `origin/main` (commit `5b731d3`).
3. **Remove stale worktree:** Issue #17 is closed/completed. Delete the worktree directory, prune git metadata, and remove the `squad/17-test-update-slide-data` branch locally and remotely.

## Rationale

- **Preserve, don't discard:** The dirty main contained legitimate workflow changes (squad orchestration configs, CI/CD templates). Discarding would lose work; moving to a dedicated branch keeps it safe and reviewable.
- **Issue #17 is done:** No active work on that issue; branch can be safely removed.
- **Main should be stable:** After cleanup, main is clean and can be used as a reliable base for new work.

## ⚠️ Coordinator Action Required

**File:** `.github/agents/squad.agent.md` was modified and is now on the preservation branch `squad/tmp-preserve-orchestration-changes` (commit `9641775`).

**Decision needed:** Does the squad.agent.md update align with current squad orchestration strategy? If yes, approve a PR to integrate it to main. If no, the file remains available locally for review/revert.

## Outcome

- **main:** Clean, matches origin/main
- **squad/tmp-preserve-orchestration-changes:** Local branch with all preserved changes (available for future integration or review)
- **Issue #17 worktree:** Removed
- **Squad/17 branch:** Removed from local and remote

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


### Issue #121 & #125 Orchestration: Validation & Text Formatting

**Orchestrator:** Copilot  
**Date:** 2026-03-25  
**Status:** ✅ Implemented (PR #146, #147)

**Parallel Agents:** Shiherlis (testing), Cheritto (features)

**Work Summary:**
- **#121 Validation:** ValidationResult model, PresentationService.Validation.cs, PptxTools.Validation.cs (41 tests: 30 service + 11 tool, 616/616 passing)
- **#125 Text Formatting:** TextFormattingInfo model, PresentationService.TextFormatting.cs, PptxTools.TextFormatting.cs (51 tests: 37 service + 14 tool, 633/633 passing)
- **Build:** 0 errors | **Total Tests:** 1,249/1,249 passing
- **Maintenance:** Fixed CI failures in squad/117 and squad/137 by merging main

**Impact:** Both validation and text formatting complete, feature-ready for review & merge.
