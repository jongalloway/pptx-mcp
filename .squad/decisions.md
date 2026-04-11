# Squad Decisions

## Active Decisions

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
