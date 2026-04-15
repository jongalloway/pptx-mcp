### Decision: Close #94 (CLI Research) — Fully Satisfied

**Owner:** McCauley (Lead)
**Date:** 2026-04-11
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

### Decision: Promote #138 (Docs) — Expand Scope, Prioritize CLI Docs

**Owner:** McCauley (Lead)
**Date:** 2026-04-11
**Status:** 🟡 Recommend `go:yes` (currently labeled `go:no`)
**Priority:** Medium

**Context:** Issue #138 covers 4 documentation gaps: EMU Calculator Guide, Shape Resolution Guide, Troubleshooting Guide, Contributing Guide. Currently labeled `go:no`.

**New finding:** The CLI (7 command groups, dual-mode architecture) shipped via #98–#104 but has **zero user-facing documentation**. The README only covers MCP server setup. A user who discovers `pptx-tools` today has no way to know they can run `pptx-tools analyze file-size presentation.pptx` from the command line.

**Decision:**

1. **Change label** from `go:no` to `go:yes` — documentation gaps are now actively hurting usability.

2. **Expand scope** to include CLI usage documentation (or create a separate issue if preferred):
   - CLI quick-start: how to run commands, dual-mode behavior
   - Command reference: all 7 command groups with examples
   - README update: add CLI section alongside MCP setup

3. **Recommended priority order:**

   | Priority | Doc | Rationale |
   |----------|-----|-----------|
   | P0 | **CLI Usage / README update** | Shipped feature with zero docs — biggest gap |
   | P1 | **Contributing Guide** | Lowers barrier for new contributors |
   | P2 | **Troubleshooting Guide** | Reduces support burden |
   | P3 | **EMU Calculator Guide** | Useful but niche |
   | P3 | **Shape Resolution Guide** | Useful but niche |

4. **Assignment:** @copilot (Coding Agent) — good fit per original triage. CLI docs can reference existing command files in `src/PptxTools/Commands/`.

**Impact:** Promoting #138 closes the documentation gap created by the CLI implementation wave (#98–#104). Without this, we shipped a feature nobody knows about.
