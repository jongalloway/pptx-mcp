# Project Context

- **Owner:** Jon Galloway
- **Project:** pptx-mcp — .NET 10 MCP server for PowerPoint manipulation via OpenXML SDK
- **Stack:** .NET 10, C#, ModelContextProtocol v1.1.0, DocumentFormat.OpenXml v3.3.0, xUnit v3 (MTP), Microsoft.Extensions.Hosting v10.0.5
- **Test project:** tests/PptxMcp.Tests/ — xUnit v3 on Microsoft Testing Platform
- **Test command:** `dotnet test --solution PptxMcp.slnx --configuration Release --no-build` (uses `--filter-method` not `--filter`)
- **Reference repos:** jongalloway/dotnet-mcp (test patterns), jongalloway/MarpToPptx (OpenXML test fixtures)
- **Created:** 2026-03-16

## Learnings

### Phase 1 E2E Testing Assignment (2026-03-16)
- Assigned #8 (E2E test: read real presentation and export markdown)
- Depends on Cheritto completing #6 & #7
- Test scope: 3+ diverse real-world presentations with accuracy validation
- Integration tests must ensure CI passes and PowerPoint compatibility verified
- Monitor Cheritto's progress on tool implementations before starting E2E suite
### Phase 2 Assignments (2026-03-16)
- **Issue #17 (shiherlis assigned):** Test pptx_update_slide_data with real metric slides — validates PowerPoint compatibility and edge cases for Goal 2A
- **Issue #15 (shiherlis assigned):** E2E test multi-source update scenario — validates full composition workflow (Goal 2B)
- Dependency: Both #17 and #15 depend on #19 (core tool implementation by cheritto) being complete
- Timeline: Phase 2 estimated 3–4 weeks after Phase 1 stabilization
- Test approach: Use TestPptxHelper.cs fixtures for realistic metric slides and multi-source composition patterns

### Phase 1 E2E Coverage Added (2026-03-16)
- TestPptxHelper now supports real speaker notes via `SpeakerNotesText`, so fixtures can validate note-aware scenarios without checking in binary decks.
- Phase 1 E2E coverage uses three generated presentations: product-update, visual-edge-cases, and unicode/localization.
- Both `pptx_extract_talking_points` and `pptx_export_markdown` are now exercised against multi-slide decks with bullets, tables, images, empty slides, image-only slides, Unicode text, and speaker notes that must stay out of Phase 1 outputs.

<!-- Append new learnings below. Each entry is something lasting about the project. -->

### Phase 4 Wave 1 Testing (2026-03-24)
- **#80 Tests:** 23 tests passing, covers file size analysis edge cases, compressed/uncompressed files, single/multi-part breakdown
- **#81 Tests:** 29 tests — discovered & fixed stream disposal bug in media analysis loop (resource leak), tests pushed to branch
- **#82 Tests:** 40 tests passing, validates layout detection, unused masters/layouts, relationship integrity
- **Total Test Coverage:** 92 new tests, all passing, Wave 1 tooling fully validated
- **Pattern:** Analysis tool testing emphasizes edge cases (empty presentations, corrupted packages, large media counts) + real-file round-trip validation
- **Quality:** Stream bug discovery demonstrates value of comprehensive test coverage; PR-ready with zero regressions

### Phase 2 Completion (2026-03-16)
- **Issues #17 & #15:** Completed and merged (PR #32 & #31)
- **Testing scope:** Issue #17 (tool testing) + Issue #15 (E2E scenario)
- **Test cases:** 7 integration tests (edge cases, format preservation, Unicode)
- **E2E scenario:** 4-slide KPI dashboard, dual targeting (shapeName + placeholderIndex), format fidelity verification
- **Quality:** Realistic fixtures (TestPptxHelper), OpenXML Validator zero errors, PowerPoint round-trip verified
- **Coverage:** 66/66 tests passing (up from 52), includes speaker notes integrity check
- **Dependency satisfaction:** Both issues unblocked by #19 (Cheritto's tool) and #18 (Copilot's docs)
- **Result:** Phase 2 testing complete, validates PowerPoint compatibility and multi-source composition pattern

### Issue #36 Table Tools Test Suite (2026-03-17)
- **Scope:** 28 new tests across 2 files for pptx_insert_table and pptx_update_table
- **Files created:**
  - `tests/PptxMcp.Tests/Services/TableOperationTests.cs` — 22 service-level tests
  - `tests/PptxMcp.Tests/Tools/TableToolsTests.cs` — 6 tool-level tests
- **Coverage:** 214/214 tests passing (up from 186)
- **Key patterns learned:**
  - Test fixtures created with TestPptxHelper produce pre-existing SlideMaster validation errors — always use baseline comparison pattern (`ValidatePresentation` before/after), never `Assert.Empty` on validator output
  - Table implementation skips out-of-range cell coordinates: returns `Success=true` with `CellsSkipped` incremented — not `Success=false`
  - `TableCellUpdate` record: `(int Row, int Column, string Value)` — 0-based coordinates
  - `InsertTable` service signature: `(filePath, slideNumber, headers[], rows[][], tableName?, x?, y?, width?, height?)` returns `TableInsertResult`
  - `UpdateTable` service signature: `(filePath, slideNumber, updates[], tableName?, tableIndex?)` returns `TableUpdateResult`
  - Table name lookup is case-insensitive; tableIndex is 0-based among tables on slide
  - GraphicData URI for tables: `http://schemas.openxmlformats.org/drawingml/2006/table` — must be exact
- **Edge cases tested:** 1x1 table, empty table (headers only), large table (13×6), custom/default positioning, unique shape IDs, cell property preservation, existing shape preservation, multiple cell updates in one call

### PR #74 Rebase — Assertion Pattern Standardization (2026-03-18)
- **Task:** Rebased `squad/65-assertion-patterns` onto `origin/main` after PRs #59, #71, #73 were merged
- **Conflicts found (3 files, 6 conflict regions):**
  - `PresentationServiceTests.cs` (3 conflicts): PptxTestBase extraction (#71) renamed `CreateTempPptx` → `CreateMinimalPptx` and `_service` → `Service`; PR #65 changed `FirstOrDefault` + `Assert.NotNull` → `Assert.Single`
  - `TemplateSlideTests.cs` (2 conflicts): Same naming changes from #71; PR #65 changed `.Single(predicate).Text` → `Assert.Single` + named variable
  - `TableToolsTests.cs` (1 conflict): Same naming changes from #71; PR #65 changed `.First(predicate)` → `Assert.Single`
- **Resolution strategy:** Keep main's naming conventions (`CreateMinimalPptx`, `Service.`) AND PR's assertion pattern improvements (`Assert.Single` over `FirstOrDefault`/`.Single`/`.First`)
- **Result:** 377/377 tests passing, CI green, force-pushed with `--force-with-lease`
### Issue #56 Theory Parameterization Refactor (2026-03-17)
- **Scope:** Consolidated 14 repetitive file-not-found `[Fact]` tests into 4 `[Theory]` parameterized tests across 3 files
- **Files modified:**
  - `tests/PptxMcp.Tests/Tools/PptxToolsTests.cs` — 8 simple "Error:" tests → 1 Theory; 2 replace_image tests → 1 Theory
  - `tests/PptxMcp.Tests/Tools/TableToolsTests.cs` — 2 table file-not-found tests → 1 Theory (using JsonDocument for type-agnostic assertion)
  - `tests/PptxMcp.Tests/Tools/ImageReplaceToolTests.cs` — 2 pptx/image-not-found tests → 1 Theory with bool flags for setup variation
- **Net result:** -55 lines, 260/260 tests passing (unchanged count), adding new tools needs only a new `[InlineData]` row
- **Patterns used:**
  - `switch` expression to dispatch tool calls by name string in Theory body
  - Bool parameters (`pptxExists`, `imageExists`) to vary test setup within a single Theory
  - `JsonDocument` for type-agnostic Success/Message assertions when result types differ
- **Key decision:** Left `PptxExportMarkdownToolTests.cs` alone (only 1 file-not-found test — no consolidation value) and kept `BothFilesMissing` test as separate `[Fact]` (tests ordering concern, not just error detection)

### Issue #67 PptxTestBase Extraction (2026-03-17)
- **Scope:** Extracted `PptxTestBase` abstract base class to eliminate duplicated setup across 16 test files
- **File created:** `tests/PptxMcp.Tests/PptxTestBase.cs`
- **Files modified:** 16 test files + `TestPptxHelper.cs` (visibility change for definition types)
- **Net result:** -239 lines, 260/260 tests passing (unchanged count), PR #71
- **Base class provides:**
  - `Service` (PresentationService), `CreateMinimalPptx()`, `CreatePptxWithSlides()`, `TrackTempFile()`, `Dispose()` with ordered cleanup
- **Key patterns:**
  - Test definition types (`TestSlideDefinition`, etc.) changed from `internal` to `public` to allow `protected` base class methods
  - Classes with custom OpenXML fixture creation (ImageReplaceTests, ImageReplaceToolTests) keep their own `CreatePptxWithPicture` and use `CreateTrackedPath()` wrapper for path generation + tracking
  - `SlideOrganizationTests` uses `CreateNamedSlides(params string[])` wrapper around `CreatePptxWithSlides`
  - `PptxCompletionHandlerTests` wraps both `CreateMinimalPptx()` and `CreatePptxWithSlides()` in a single `CreateTempPptx(params TestSlideDefinition[])` for backward compatibility
  - `PptxPromptsTests` has no disposable pattern — correctly excluded from base class

### Issue #80 File Size Analysis Test Suite (2026-03-18)
- **Scope:** 23 new tests in `tests/PptxMcp.Tests/Services/FileSizeAnalysisTests.cs` for `AnalyzeFileSize` service method
- **Written proactively** while Cheritto implements `PresentationService.Optimization.cs` — aligned tests to WIP model/service signatures
- **Categories tested:** happy path (minimal/multi-slide), image categorization, masters/layouts separation, empty categories (0 not null), arithmetic invariants (subtotals sum to grand total), metadata quality (non-empty paths/content types/non-negative sizes), error handling (file not found), complex fixture (tables + charts + images combined)
- **Key findings:**
  - OpenXML chart definitions may add extra parts under `/ppt/slides/` — use `>= N` assertions for slide count in complex fixtures, not exact equality
  - TestPptxHelper with `IncludeImage = true` adds PNG `ImagePart` per slide — each becomes a separate part in the "images" category
  - `CategorizePart` logic excludes `.rels` files from content categories — they fall to "other"
  - `FileSizeAnalysisResult` has both `TotalFileSize` (disk) and `TotalPartSize` (uncompressed sum) — distinct values
- **Test count:** 441/441 passing (up from 418)

### Issue #121 Validation & Diagnostics Test Suite (2026-03-25)
- **Scope:** 41 new tests across 2 files for pptx_validate_presentation
- **Files created:**
  - `tests/PptxTools.Tests/Services/ValidationTests.cs` — 30 service-level tests
  - `tests/PptxTools.Tests/Tools/ValidationToolsTests.cs` — 11 tool-level tests
- **Coverage:** 616/616 tests passing (up from 575)
- **Key patterns learned:**
  - Tool was initially named `pxtx_validate_presentation` (typo), Cheritto fixed to `pptx_validate_presentation` in follow-up commit
  - Validation checks: DuplicateShapeId, MissingImageReference, MissingRequiredElement, OrphanedRelationship, BrokenHyperlinkTarget, CrossSlideDuplicateShapeId
  - `ValidatePresentation(filePath, slideNumber?)` — slideNumber filter skips cross-slide check
  - Cross-slide shape ID duplicates are Info severity (common in normal presentations)
  - TestPptxHelper reuses shape IDs starting at 2 per slide, so multi-slide fixtures naturally produce CrossSlideDuplicateShapeId issues
  - Corrupt fixtures created by opening valid PPTX with `PresentationDocument.Open(path, true)` and modifying XML directly
  - Concurrent agents on same repo can switch branches — must verify branch before each operation
- **Edge cases tested:** duplicate shape IDs (within-slide and cross-slide), missing image ref (broken blip embed rId999), missing CommonSlideData, missing ShapeTree, empty presentation (no slides), file not found, slide number filtering, severity sorting, idempotency
