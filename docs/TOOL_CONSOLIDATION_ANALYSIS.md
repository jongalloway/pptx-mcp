# Tool Consolidation Analysis

**Issue:** #92 — Q4: Tool consolidation analysis
**Author:** McCauley (Lead)
**Date:** 2026-03-25
**Status:** Implemented (Tier 1 + update_text deprecation — PR for #96)

## Executive Summary

We have **24 MCP tools** after Phase 4. This analysis evaluates consolidation opportunities to reduce tool count and improve LLM tool selection. The recommendation is to consolidate **4 tools into 2** (Tier 1), with **2 additional reductions** as stretch goals (Tier 2), bringing the total from **24 → 20 tools**.

The guiding principle: only consolidate when the resulting tool is **clearer and more natural** to use than the individual tools.

---

## Current Tool Inventory (24 tools)

### Already Consolidated (3 multi-action tools)

| Tool | Actions | File |
|------|---------|------|
| `pptx_manage_slides` | Add / AddFromLayout / Duplicate | PptxTools.ManageSlides.cs |
| `pptx_reorder_slides` | Move / Reorder | PptxTools.ReorderSlides.cs |
| `pptx_chart_data` | Read / Update | PptxTools.Charts.cs |

### Standalone Tools (21 tools)

| # | Tool | Category | Read/Write | File |
|---|------|----------|------------|------|
| 1 | `pptx_list_slides` | Navigation | Read | PptxTools.cs |
| 2 | `pptx_list_layouts` | Navigation | Read | PptxTools.cs |
| 3 | `pptx_get_slide_xml` | Inspection | Read | PptxTools.cs |
| 4 | `pptx_get_slide_content` | Inspection | Read | PptxTools.cs |
| 5 | `pptx_extract_talking_points` | Inspection | Read | PptxTools.cs |
| 6 | `pptx_export_markdown` | Export | Write | PptxTools.cs |
| 7 | `pptx_update_text` | Editing | Write | PptxTools.cs |
| 8 | `pptx_update_slide_data` | Editing | Write | PptxTools.cs |
| 9 | `pptx_batch_update` | Editing | Write | PptxTools.cs |
| 10 | `pptx_write_notes` | Editing | Write | PptxTools.cs |
| 11 | `pptx_insert_image` | Editing | Write | PptxTools.cs |
| 12 | `pptx_replace_image` | Editing | Write | PptxTools.cs |
| 13 | `pptx_insert_table` | Editing | Write | PptxTools.cs |
| 14 | `pptx_update_table` | Editing | Write | PptxTools.cs |
| 15 | `pptx_delete_slide` | Editing | Write | PptxTools.cs |
| 16 | `pptx_analyze_file_size` | Optimization | Read | PptxTools.Optimization.cs |
| 17 | `pptx_find_unused_layouts` | Optimization | Read | PptxTools.Optimization.cs |
| 18 | `pptx_remove_unused_layouts` | Optimization | Write | PptxTools.Optimization.cs |
| 19 | `pptx_optimize_images` | Optimization | Write | PptxTools.Optimization.cs |
| 20 | `pptx_analyze_media` | Media | Read | PptxTools.Media.cs |
| 21 | `pptx_deduplicate_media` | Media | Write | PptxTools.Deduplication.cs |

---

## Existing Consolidation Pattern

All three existing consolidated tools follow the same pattern:

1. **C# enum** for the `action` parameter (e.g., `ManageSlidesAction`)
2. **All action-specific params are nullable** with per-action validation
3. **`[McpMeta]` attributes**: `consolidatedTool = true` and `actions` JSON array
4. **Switch expression** dispatches to correct implementation
5. **Clean break**: old tools removed, new tool replaces them in one PR

This pattern works well when actions **share most parameters** and the consolidated name clearly communicates the tool's domain.

---

## Consolidation Candidates

### Tier 1: Easy Wins (Recommended)

#### 1. Layout Optimization → `pptx_manage_layouts`

**Current tools:**
- `pptx_find_unused_layouts(filePath)` — read-only, identifies unused layouts/masters with space impact
- `pptx_remove_unused_layouts(filePath, layoutUris?)` — write, removes unused layouts, validates package

**Proposed consolidated tool:**
```
pptx_manage_layouts(
    filePath: string,
    action: ManageLayoutsAction,    // Find | Remove
    layoutUris?: string[]           // Only for Remove; omit for auto-detect
)
```

**Why this works:**
- Both tools operate on the same domain (layout cleanup)
- Natural workflow: Find (diagnostic) → Remove (action)
- High parameter overlap — `layoutUris` is optional and action-specific
- Follows the exact pattern of `pptx_chart_data` (Read/Update)
- LLM benefit: single tool name for the entire layout cleanup workflow

**Test impact:**
- Service tests (UnusedLayoutsTests: 40, RemoveLayoutsTests: 17): **Unchanged** — service layer untouched
- Tool tests: Need action param added; ~5 methods in PptxToolsTests to update
- **Estimated effort:** 2–3 hours

**Risk:** Low. Clean action split between read and write.

**Saves:** 1 tool (2 → 1)

---

#### 2. Media Operations → `pptx_manage_media`

**Current tools:**
- `pptx_analyze_media(filePath)` — read-only, inventories media assets, detects duplicates by SHA256
- `pptx_deduplicate_media(filePath)` — write, consolidates identical media, removes orphans

**Proposed consolidated tool:**
```
pptx_manage_media(
    filePath: string,
    action: ManageMediaAction       // Analyze | Deduplicate
)
```

**Why this works:**
- **Perfect** parameter overlap — both take only `filePath`
- Natural workflow: Analyze (identify duplicates) → Deduplicate (clean them up)
- Same analyze→act pattern as layout tools above
- Simplest possible consolidation — no action-specific params needed

**Test impact:**
- Service tests (MediaAnalysisTests: 34, DeduplicateMediaTests: 10): **Unchanged**
- Tool tests: ~3 methods in PptxToolsTests to update
- **Estimated effort:** 2 hours

**Risk:** Low. Identical parameter surfaces, clean action split.

**Saves:** 1 tool (2 → 1)

---

### Tier 2: Moderate (Stretch Goals)

#### 3. Slide Inspection → `pptx_get_slide`

**Current tools:**
- `pptx_get_slide_xml(filePath, slideIndex)` — returns raw OpenXML
- `pptx_get_slide_content(filePath, slideIndex)` — returns structured JSON (shapes, text, positions)

**Proposed consolidated tool:**
```
pptx_get_slide(
    filePath: string,
    slideIndex: int,
    format: SlideFormat              // Xml | Content
)
```

**Why this is moderate, not easy:**
- Parameters are identical (good) but the tools serve very different audiences
- `get_slide_xml` is a debugging/deep-inspection tool; `get_slide_content` is the primary structured reader
- An LLM might actually select *more accurately* with descriptive names than with a `format` enum
- However, having two "get slide" tools with nearly identical names already causes selection ambiguity

**Test impact:**
- `get_slide_content` is referenced in **72 places** across 4+ test files (PresentationServiceTests, BoundaryConditionTests, UpdateSlideDataTests, PptxToolsTests, TableOperationTests)
- `get_slide_xml` referenced in ~18 places
- **Estimated effort:** 4–5 hours (widespread test updates)

**Risk:** Medium. High test surface. Debatable LLM clarity improvement.

**Recommendation:** Defer unless Tier 1 consolidations go smoothly and team has bandwidth.

**Saves:** 1 tool (2 → 1)

---

#### 4. Deprecate `pptx_update_text`

**Current tools:**
- `pptx_update_text(filePath, slideIndex, placeholderIndex, text)` — Phase 1 text update
- `pptx_update_slide_data(filePath, slideNumber, shapeName?, placeholderIndex?, newText)` — Phase 2 superset

**Analysis:**
`pptx_update_slide_data` is a **strict functional superset** of `pptx_update_text`. It adds shape name targeting (preferred) while retaining placeholder index fallback. Having both confuses LLMs about which to use.

This isn't a consolidation — it's a **deprecation** of a redundant tool.

**Test impact:**
- Only 1–2 direct tool tests for `update_text`
- No service-layer impact (different service method)
- **Estimated effort:** 1 hour

**Risk:** Low, but check external documentation/examples referencing `update_text`.

**Saves:** 1 tool

---

### Tier 3: Don't Consolidate

#### 5. Analysis Mega-Tool (`pptx_analyze` with target param)

**Issue #92 suggests:** `analyze_file_size` + `analyze_media` + `find_unused_layouts` → `pptx_analyze(target: FileSize/Media/Layouts)`

**Why this is a bad idea:**
1. **Conflicts with better pairings.** If we consolidate `analyze_media` with `deduplicate_media` (Tier 1) and `find_unused_layouts` with `remove_unused_layouts` (Tier 1), those tools are already spoken for. The mega-tool would orphan the corresponding write operations.
2. **Breaks workflow coherence.** Domain pairings (analyze→act) are more natural for agents than grouping all diagnostics together. An agent cleaning up media wants `manage_media(Analyze)` → `manage_media(Deduplicate)`, not `pptx_analyze(Media)` → `pptx_deduplicate_media()`.
3. **`analyze_file_size` stands alone fine.** It's the only global file-level diagnostic — no corresponding write action to pair with.

**Verdict:** Don't consolidate. Domain-specific pairings are stronger.

---

#### 6. Image Tools (`insert_image` + `replace_image`)

**Why not:**
- Very different parameter surfaces: `insert_image` takes position/size EMUs; `replace_image` takes shape targeting (name/index)
- Different intent: creating new vs replacing existing
- Consolidation would require many nullable params with confusing per-action rules

**Verdict:** Don't consolidate. Parameter divergence too high.

---

#### 7. Table Tools (`insert_table` + `update_table`)

**Why not:**
- Same argument as images: creating (headers/rows arrays) vs updating (cell update array)
- Different parameter shapes with minimal overlap
- `pptx_chart_data` works because Read and Update share the same chart targeting params

**Verdict:** Don't consolidate.

---

#### 8. `pptx_optimize_images` (standalone)

Unique parameter surface (`targetDpi`, `jpegQuality`, `convertFormats`). No natural pairing candidate. Leave standalone.

---

## Tool Count Projection

| Phase | Tool Count | Change |
|-------|-----------|--------|
| Current | 24 | — |
| After Tier 1 (layouts + media) | **22** | -2 |
| After deprecate update_text | **21** | -1 |
| **Implemented (PR #96)** | **21** | **-12.5%** |
| After Tier 2 (slide inspection, if desired) | **20** | -1 |

---

## Implementation Notes

**Implemented 2026-03-26 by Cheritto (PR for issue #96):**

### What shipped
1. **`pptx_manage_layouts`** (Find | Remove) — replaced `pptx_find_unused_layouts` + `pptx_remove_unused_layouts`
2. **`pptx_manage_media`** (Analyze | Deduplicate) — replaced `pptx_analyze_media` + `pptx_deduplicate_media`
3. **Deprecated `pptx_update_text`** — `pptx_update_slide_data` is a strict functional superset

### Pattern followed
All consolidated tools use the same pattern as `pptx_manage_slides`, `pptx_reorder_slides`, and `pptx_chart_data`:
- C# enum for action parameter
- `[McpMeta]` attributes for machine-readable metadata
- `partial` method, switch expression dispatch
- Service layer unchanged

### Test impact
- 552/552 tests passing (unchanged count)
- Service-layer tests (UnusedLayoutsTests, RemoveLayoutsTests, MediaAnalysisTests, DeduplicateMediaTests, ImageOptimizationTests) all pass unchanged
- Tool-level test for `pptx_update_text` removed; null validation test replaced with `pptx_manage_layouts` and `pptx_manage_media` coverage

### Files changed
- `PptxTools.Optimization.cs` — consolidated layout tools into `pptx_manage_layouts`
- `PptxTools.ManageMedia.cs` — new file, consolidated media tools into `pptx_manage_media`
- `PptxTools.Media.cs` — deleted (absorbed into ManageMedia)
- `PptxTools.Deduplication.cs` — deleted (absorbed into ManageMedia)
- `PptxTools.cs` — removed `pptx_update_text` method
- `ManageLayoutsAction.cs`, `ManageMediaAction.cs` — new enum files
- `PptxToolsTests.cs`, `NullValidationTests.cs` — updated for new tool names
- `README.md` — tool list updated (24 → 21)

---

## Recommended Implementation Plan

### Phase 1: Tier 1 Consolidations (1 PR)

**Scope:** Layout and media tool consolidation

1. Create `ManageLayoutsAction` enum (`Find` / `Remove`)
2. Create `ManageMediaAction` enum (`Analyze` / `Deduplicate`)
3. Implement `pptx_manage_layouts` in `PptxTools.Optimization.cs` (replaces both layout tools)
4. Implement `pptx_manage_media` in new `PptxTools.ManageMedia.cs` (replaces analyze + dedup)
5. Delete old tool methods
6. Update tool-level tests (service tests unchanged)
7. Update README.md and TOOL_REFERENCE.md

**Migration path:** Clean break in single PR (same pattern as manage_slides consolidation). No deprecation period — the old tool names never shipped as a stable API.

**Estimated effort:** 4–5 hours

### Phase 2: Tier 2 Reductions (separate PRs)

**PR A:** Deprecate `pptx_update_text` (1 hour)
**PR B:** Consolidate slide inspection tools (4–5 hours, if desired)

Each as a separate PR for clean review.

---

## Decision Criteria

A consolidation is recommended when ALL of these hold:

1. **Shared domain:** Tools operate on the same conceptual entity
2. **Parameter overlap ≥ 80%:** Most params are shared; action-specific params are few and optional
3. **Natural workflow:** Actions represent a logical progression (analyze → act)
4. **Clearer for LLMs:** The consolidated tool name + action enum is at least as easy to select as the individual tools
5. **Manageable test impact:** Service-layer tests survive unchanged; only tool-level tests need updates

A consolidation is **rejected** when:

- Parameter surfaces diverge significantly (image tools, table tools)
- It conflicts with a stronger domain pairing (analysis mega-tool)
- Individual tool names are already maximally clear (debatable for slide inspection)

---

## Documentation Note

**TOOL_REFERENCE.md is stale.** It still lists pre-consolidation tools (`pptx_add_slide`, `pptx_add_slide_from_layout`, `pptx_duplicate_slide`, `pptx_move_slide`) that were replaced by `pptx_manage_slides` and `pptx_reorder_slides`. This should be fixed alongside any consolidation work.
