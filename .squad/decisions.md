# Squad Decisions

## Active Decisions

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

### Feasibility Decision: Issues #115 & #116 — Shape Creation & Animation Support

**Author:** Nate (Consulting Dev)  
**Date:** 2026-03-17T06:15Z  
**Related Issues:** #115 (Shape creation), #116 (Animation/transitions)  
**Stakeholders:** Jon (Owner), McCauley (Architect), Cheritto (Backend Dev)  
**Status:** READY FOR GREENLIGHTING

---

## Executive Summary

Two high-complexity features have been thoroughly researched and are ready for prioritization:

| Issue | Feature | Verdict | Effort | Timeline |
|-------|---------|---------|--------|----------|
| **#115** | Shape creation (rectangles, circles, arrows, connectors) | 🟢 GO:YES | L–XL | 5–8 weeks (MVP) |
| **#116a** | Slide transitions (6 types) | 🟢 GO:YES | L–M | 1–2 weeks |
| **#116b** | Shape animations (entrance/exit/emphasis) | 🟡 GO:CONDITIONAL | H–XL | 4–6 weeks (Phase 2) |

**Key Finding:** Both shape creation and transitions have **excellent OpenXML support** and proven patterns from MarpToPptx. Animations are deliberately avoided in MarpToPptx due to complexity; recommend same for pptx-tools.

---

## Decision 1: Issue #115 — Shape Creation

### Recommendation: 🟢 GO:YES — Promote to Development

### Rationale

1. **SDK Support: Excellent**
   - DocumentFormat.OpenXml v3.5.1 has first-class Shape, Picture, PresetGeometry classes
   - 100+ preset geometries available (Rectangle, Ellipse, Arrows, Connectors, Flowchart, Callouts)
   - API is stable and well-documented

2. **Prior Art Proven**
   - MarpToPptx implements shape creation in `/src/MarpToPptx.Pptx/Rendering/OpenXmlPptxRenderer.cs`
   - Patterns are straightforward and maintainable
   - Code is reference-quality for pptx-tools implementation

3. **pptx-tools Fit**
   - Service architecture (partial classes) aligns perfectly
   - Existing patterns (safe null navigation, shape tree iteration) apply directly
   - No architectural blockers identified

4. **MVP Scope Clear**
   - **Phase 1a (2–3 weeks):** Core shapes — Rectangle, Ellipse, Diamond, Triangle, Pentagon, Hexagon, Star5
   - **Phase 1b (1 week):** Arrows & connectors — 4 directions, bidirectional, bent/curved/elbow variants
   - **Phase 1c (2 weeks):** Read/update properties — fill, stroke, position, size
   - Total MVP: 5–8 weeks, achievable in one development cycle

5. **Business Value**
   - AI-driven diagram creation (major use case)
   - Layout customization and flowcharts (high demand)
   - Direct support for presentation automation workflows

6. **Risk Level: Low–Medium**
   - Shape ID collisions: Low risk (max + 1 lookup pattern is proven)
   - EMU precision: Medium risk (unit conversion helper function mitigates)
   - Property ordering: Low risk (schema enforces order; SDK validates)
   - Theme inheritance: Medium risk (defer to Phase 2; start with explicit RGB)

### Implementation Path

**Recommended Workflow:**
1. Create `PresentationService.Shapes.cs` partial class
2. Implement `CreateShape()`, `GetShapeProperties()`, `UpdateShapeProperties()` methods
3. Create MCP tool wrapper: `pptx_shape_operations` (actions: Create, Read, Update)
4. Follow pptx-tools conventions (partial class, result objects, safe null navigation)
5. Mirror MarpToPptx patterns exactly (NonVisualShapeProperties + ShapeProperties structure)

**Assignment Recommendation:** Cheritto (Backend Dev) — strong fit for OpenXML work

**Delivery Timeline:** 5–8 weeks for full MVP

---

## Decision 2: Issue #116 — Animation & Transition Support

### Sub-Decision 2a: Transitions (Slide Transition Effects) — 🟢 GO:YES

### Rationale

1. **SDK Support: Excellent**
   - Transition classes well-defined in OpenXML (Fade, Cut, Push, Wipe, Cover, Pull, etc.)
   - Simple enum-based API with clear direction/speed/duration parameters
   - Proven by MarpToPptx (`ApplyTransition()` method, lines ~4500–4800)

2. **Prior Art Proven**
   - MarpToPptx implements transitions directly and maintainably
   - Pattern is straightforward: build Transition element → insert after CommonSlideData/ColorMapOverride
   - Direction mapping is simple enum lookup

3. **Business Value**
   - Dynamic presentations (professional polish)
   - Automated slideshow design (high-demand for AI workflows)
   - Low complexity, immediate ROI

4. **MVP Scope Clear**
   - **pptx_get_slide_transitions:** Read transition type, speed, duration, direction
   - **pptx_set_slide_transition:** Write 6 types (Fade, Cut, Push, Wipe, Cover, Pull), speed, direction
   - Total MVP: 1–2 weeks

5. **Risk Level: Low**
   - Enum-based API with clear validation
   - MarpToPptx precedent proves reliability
   - No complex state management or XML tree corruption risk

### Sub-Decision 2b: Shape Animations (Entrance/Exit/Emphasis) — 🟡 GO:CONDITIONAL

### Rationale for Deferral

1. **SDK Support: Weak**
   - Animation model uses deeply nested CTTiming tree (ParallelTimeNode, SequenceTimeNode, Animate classes)
   - No preset shortcuts; most animations require manual XML construction
   - Poor documentation; low usage in practice

2. **Prior Art: Deliberately Avoided**
   - MarpToPptx **does not implement** shape animations
   - Only transitions implemented (simpler model)
   - Team choice made consciously due to complexity

3. **Complexity vs. Demand Mismatch**
   - Estimated effort: 4–6 weeks for basic animations (Appear, Fade, Fly-in, Zoom)
   - Most presentations don't use animations (design trend favors static slides)
   - User demand unvalidated

4. **Testing Challenge**
   - Animations are difficult to test without running PowerPoint interactively
   - Timing tree corruption is subtle and hard to debug
   - Risk of shipping broken animations

### Conditional Approval Criteria

**Approve Phase 2 (Animations) IF AND ONLY IF:**

1. ✅ Transitions (Phase 1) ship successfully and receive positive user feedback
2. ✅ Multiple GitHub issues request animation support (validate demand)
3. ✅ McCauley approves as architectural priority (not just nice-to-have)
4. ✅ Resources available for 4–6 week commitment + spike investigation

**If conditions not met by Month 2 (May 2026):** Defer indefinitely as low-priority nice-to-have.

### Implementation Path (When Approved)

1. **Spike Investigation (1 week):**
   - Prototype basic entrance animations (Appear, Fade, Fly-in) on simple test cases
   - Identify gotchas and XML validation issues
   - Document timing model and event trigger patterns

2. **Phase 2 Implementation (3–4 weeks):**
   - `PresentationService.Animations.cs` partial class
   - `AddAnimation()`, `GetAnimations()` methods
   - MCP tool: `pptx_animation_operations` (actions: Add, List)
   - Heavy unit testing (animation timing is subtle)

3. **Acceptance Criteria (High Bar):**
   - 15+ unit test cases (animation timing scenarios)
   - E2E validation on 5+ real presentations
   - Animations preview correctly in PowerPoint
   - No timing tree corruption after updates

**Assignment Recommendation:** Cheritto + Shiherlis (pair programming for complex timing logic)

---

## Implementation Priority Sequence

### Immediate (Week 1–2)

1. **Shapes Phase 1a:** Core shapes (rectangles, circles, basic shapes)
   - Issue: #115 Phase 1a
   - Assignee: Cheritto
   - 2–3 weeks effort

2. **Transitions Phase 1:** Slide transitions (6 types)
   - Issue: #116 Phase 1
   - Assignee: Cheritto (parallel with shapes Phase 1b)
   - 1–2 weeks effort

### Medium-term (Week 4–8)

3. **Shapes Phase 1b–c:** Arrows, connectors, read/update properties
   - Issue: #115 Phase 1b–c
   - Assignee: Cheritto
   - 2–3 weeks effort

### Conditional (After Validation)

4. **Animations Phase 2:** Entrance/exit/emphasis animations
   - Issue: #116 Phase 2 (create after user demand validated)
   - Assignee: Cheritto + Shiherlis (spike + implementation)
   - 4–6 weeks effort (conditional)

---

## Technical Patterns (Ready to Use)

### Shape Creation Pattern

```csharp
// From MarpToPptx reference
var shape = new P.Shape(
    new P.NonVisualShapeProperties(
        new P.NonVisualDrawingProperties { Id = nextId, Name = "Shape 1" },
        new P.ApplicationNonVisualDrawingProperties()),
    new P.ShapeProperties(
        new A.Transform2D(
            new A.Offset { X = ToEmu(x), Y = ToEmu(y) },
            new A.Extents { Cx = ToEmu(width), Cy = ToEmu(height) }),
        new A.PresetGeometry(new A.AdjustValueList()) 
        { Preset = A.ShapeTypeValues.Rectangle },
        new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })),
    new P.TextBody());

shapeTree.AppendChild(shape);
```

### Transition Creation Pattern

```csharp
// From MarpToPptx reference
var transition = new P.Transition()
{
    Speed = P.TransitionSpeedValues.Medium,
    Duration = 500
};
transition.AppendChild(new P.FadeTransition());

slide.InsertAfterSelf(transition);
```

---

## Open Questions & Notes

1. **Theme Color Support:** Shapes can inherit fill colors from slide theme. MVP uses explicit RGB; defer theme-aware colors to Phase 2.

2. **Shape Z-order:** OpenXML determines shape layering by order in ShapeTree. MVP creates shapes at end (top layer); advanced z-order control deferred.

3. **Connector Endpoints:** Smart connector routing (auto-connect to shape edges) deferred to Phase 2. MVP supports basic line connectors only.

4. **Animation Synchronization:** Multiple animations per shape require careful timing. Defer to Phase 2 if approved.

---

## Risk Mitigation Summary

| Risk | Severity | Mitigation |
|------|----------|-----------|
| **Shape ID collisions** | Low | Use `max(existing IDs) + 1` lookup; validate uniqueness |
| **EMU precision errors** | Medium | Helper function for unit conversion; test with common sizes |
| **Invalid XML structure** | Low | OpenXML SDK validates schema; let it fail fast |
| **Theme color dependency** | Medium | Start with explicit RGB; document as future work |
| **Animation tree corruption** | High | **Defer animations to Phase 2 conditional; transitions only Phase 1** |
| **Testing without PowerPoint** | Medium | E2E test suite with real .pptx files; visual inspection required |

---

## Deliverables

1. **Feasibility Report:** `FEASIBILITY_REPORTS.md` (comprehensive analysis with code examples)
2. **History Update:** `.squad/agents/nate/history.md` (learnings captured)
3. **This Decision Record:** Team-visible rationale for prioritization

---

## Next Steps

1. **Jon reviews** this decision and feasibility report
2. **Jon approves** Phase 1a (Shapes) + Phase 1 (Transitions)
3. **McCauley creates** GitHub issues for:
   - #115-Phase-1a: Core shape creation
   - #116-Phase-1: Slide transitions
4. **Cheritto begins** Phase 1a (target: 3 weeks to completion)
5. **Team validates** demand for Phase 2 (Animations) in parallel

---

**Status:** Ready for greenlighting. No blocking questions. Recommend approval.

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
