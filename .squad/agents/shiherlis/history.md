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

### Issue #123 MCP Prompt Template Tests (2026-03-24)
- **Scope:** 44 new test methods (289 lines) for comprehensive prompt template coverage in `tests/PptxTools.Tests/Prompts/PptxPromptsTests.cs`
- **Context:** Written for 4 new prompt methods being added by Cheritto (BatchUpdateFromCsv, ExtractForBlog, CreateSpeakerNotesOutline, OptimizeForWeb) + full coverage for 3 existing prompts (RefreshQbrDeck, CreateAgendaSlide, ReplaceKpiPlaceholders)
- **Pattern:** Each prompt tested for: returns ≥1 message, first message is User role, contains file path, optional params use defaults when null, expected tool references in text
- **Test breakdown:**
  - RefreshQbrDeck: 7 tests (metricsSource optional param, mentions pptx_update_slide_data/pptx_list_slides)
  - CreateAgendaSlide: 6 tests (no optional params, mentions pptx_manage_slides/pptx_list_layouts)
  - ReplaceKpiPlaceholders: 7 tests (placeholderPattern optional param, mentions pptx_get_slide_content/pptx_update_slide_data)
  - BatchUpdateFromCsv: 7 tests (csvPath required param, mentions pptx_batch_update)
  - ExtractForBlog: 9 tests (format optional param defaults to "markdown", mentions pptx_export_markdown/pptx_extract_talking_points)
  - CreateSpeakerNotesOutline: 10 tests (style optional param with switch logic for bullet-points/narrative/timing-cues, mentions pptx_write_notes)
  - OptimizeForWeb: 10 tests (targetSizeMb optional param, mentions pptx_analyze_file_size/pptx_optimize_images/pptx_manage_media/pptx_manage_layouts)
- **Key findings:**
  - Prompt methods return `IEnumerable<PromptMessage>` with single yielded message
  - All prompts use User role (not Assistant or System)
  - Content is TextContentBlock with multiline string text
  - Optional params use `string.IsNullOrWhiteSpace` or nullable types with null-coalescing logic
  - Prompts reference tool names by their MCP names (e.g., "pptx_batch_update" not "BatchUpdate")
- **Result:** Branch `squad/123-additional-mcp-prompts` pushed with tests; full prompt template test coverage achieved (7/7 prompts tested)
### Issue #114 Hyperlink Support Test Suite (2026-03-24)
- **Scope:** 49 new tests across 2 files for hyperlink CRUD operations and MCP tool
- **Files created:**
  - `tests/PptxTools.Tests/Services/HyperlinkTests.cs` — 30 service-level tests
  - `tests/PptxTools.Tests/Tools/HyperlinkToolsTests.cs` — 19 tool-level tests
- **Written proactively** while Cheritto implements `PresentationService.Hyperlinks.cs` — aligned tests to WIP model/service/tool signatures
- **Service test coverage:** GetHyperlinks (empty, external URL, mailto, slide filtering, tooltip, multiple per slide, file-not-found), AddHyperlink (happy path, discoverable via Get, tooltip, mailto, non-existent shape/slide/file), UpdateHyperlink (URL change, tooltip change, preserves others, no-hyperlink-throws, non-existent shape), RemoveHyperlink (success, verify via Get, preserves text, no-hyperlink-throws, non-existent shape), full CRUD round-trip
- **Tool test coverage:** Structured `HyperlinkResult` assertions for Get/Add/Update/Remove actions, file-not-found across all actions, missing parameter validation (slideNumber, shapeName, url)
- **Key patterns:**
  - Custom OpenXML fixture helper `CreatePptxWithHyperlinks` creates presentations with run-level `A.HyperlinkOnClick` + `HyperlinkRelationship` for external URLs
  - `FakeGrouping` helper enables empty-slide fixture creation
  - Service returns `HyperlinkResult` record (not void) — Add/Update/Remove all return structured results with Success, Action, Url, HyperlinkCount
  - `HyperlinkAction` enum (Get, Add, Update, Remove) mirrors `ChartDataAction` consolidation pattern
  - Tool uses `ExecuteToolStructured` pattern returning structured `HyperlinkResult` on both success and error paths
  - Update/Remove throw `InvalidOperationException` when shape has no hyperlink; both throw `ArgumentException` for non-existent shapes, `ArgumentOutOfRangeException` for invalid slide numbers
- **Build status:** Cheritto's WIP service has compilation errors (`RunProperties.HyperlinkOnClick` doesn't exist — need `GetFirstChild<A.HyperlinkOnClick>()`; `P.NonVisualDrawingProperties` vs `A.NonVisualDrawingProperties` type mismatch in `GetNonVisualDrawingProperties`). Tests cannot run until those are fixed.
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

### Issue #133 Compare Presentations Test Suite (2026-03-25)
- **Scope:** 62 new tests across 2 files for pptx_compare_presentations
- **Files created:**
  - `tests/PptxTools.Tests/Services/CompareTests.cs` — 35 service-level tests
  - `tests/PptxTools.Tests/Tools/CompareToolsTests.cs` — 27 tool-level tests
- **Written proactively** while Cheritto implements `PresentationService.Compare.cs` — aligned tests to expected model/service/tool signatures
- **Expected model types:** `CompareAction { Full, SlidesOnly, TextOnly, MetadataOnly }`, `ComparisonResult` (Success, Action, SourceFile, TargetFile, AreIdentical, DifferenceCount, SlideDifferences, TextDifferences, MetadataDifferences, Message), `SlideDifference` (SlideNumber, DifferenceType, Description), `TextDifference` (SlideNumber, ShapeName, SourceText, TargetText), `MetadataDifference` (Property, SourceValue, TargetValue)
- **Expected service signature:** `ComparePresentations(string sourceFilePath, string targetFilePath, CompareAction action)` returns `ComparisonResult`
- **Expected tool signature:** `pptx_compare_presentations(string sourceFilePath, string targetFilePath, CompareAction action)` returns `Task<string>`
- **Service test coverage:** identical presentations (6 tests), slide count differences (4 tests), text changes (5 tests), metadata differences (5 tests), action-specific routing for SlidesOnly/TextOnly/MetadataOnly (9 tests), error handling (3 tests), edge cases (4 tests including same-file, empty, DifferenceCount invariant, idempotency)
- **Tool test coverage:** Full/SlidesOnly/TextOnly/MetadataOnly action routing (8 tests), file-not-found for source/target/both (4 tests), Theory for all actions on missing file (1 test), JSON structure validation (6 tests including field presence, indented output, sub-object fields)
- **Key patterns:**
  - Compare tool takes TWO file paths — uses custom file checking like `pptx_replace_image` pattern
  - Metadata fixtures use `doc.PackageProperties.Title/Creator/etc.` via `PresentationDocument.Open(path, true)`
  - `CreateIdenticalPair()` helper creates two separate but structurally identical files
  - DifferenceCount invariant: must equal sum of SlideDifferences + TextDifferences + MetadataDifferences
- **Build status:** 10 expected compilation errors — all reference `CompareAction`/`ComparisonResult`/`pptx_compare_presentations` which Cheritto will create. Zero syntax errors in test code.

### Issue #128 Export JSON Test Suite (2026-03-25)
- **Scope:** 76 new tests across 2 files for pptx_export_json
- **Files created:**
  - `tests/PptxTools.Tests/Services/ExportJsonTests.cs` — 46 service-level tests
  - `tests/PptxTools.Tests/Tools/ExportJsonToolsTests.cs` — 30 tool-level tests
- **Written proactively** while Cheritto implements — aligned tests to WIP model/service/tool signatures
- **Model types:** `ExportJsonAction { Full, SlidesOnly, MetadataOnly, SchemaOnly }`, `PresentationExport` (Success, Action, FilePath, Metadata, SlideCount, Slides, Schema, Message), `SlideExport` (SlideNumber, SlideIndex, Title, SlideWidthEmu, SlideHeightEmu, Shapes, SpeakerNotes), `ShapeExport` (with embedded Table/Image/Chart optional params), `TableExportData`, `ImageExport`, `ChartExport`
- **Service test coverage:** minimal presentation (7 tests), text shapes (4 tests), tables via ShapeExport.Table (4 tests), images via ShapeExport.Image (3 tests), speaker notes (2 tests), MetadataOnly (4 tests), SlidesOnly (5 tests), SchemaOnly (5 tests), multi-slide (2 tests), mixed content (1 test), error handling (1 test), action string Theory (1 test), metadata fields (1 test), idempotency (1 test), notes per slide (1 test), shape names (1 test), invariants (2 tests), charts (2 tests), shape types (2 tests)
- **Tool test coverage:** Full/SlidesOnly/MetadataOnly/SchemaOnly action routing (11 tests), file-not-found (4 tests), null/empty file path validation (2 tests), JSON structure (7 tests including field presence, indented output, error fields, slide/metadata/table/image/chart JSON sub-structures)
- **Key findings:**
  - Cheritto's model evolved during parallel development: `SlideExport` moved from separate Tables/Images/Charts/Notes params to embedded ShapeExport sub-types + computed properties + SpeakerNotes
  - `PresentationExport` gained `Schema` field; `ChartExport`/`ImageExport` lost `ShapeName` (embedded in ShapeExport)
  - **Bug discovered:** `ExtractSlideTitle` in `PresentationService.ExportJson.cs` checks `PlaceholderType is "title" or "ctrTitle"` but `ShapeContent.PlaceholderType` stores the C# enum name (`"Title"` / `"CenteredTitle"`), not the XML attribute value — causes 2 test failures (title always null)
  - SchemaOnly action bypasses file I/O entirely — `Service.ExportJson("", SchemaOnly)` is valid
  - Tool parameter `filePath` is `string?` (nullable) — SchemaOnly works with null, other actions return structured error
- **Build status:** 0 compilation errors, 74/76 tests passing, 2 failures from title extraction bug (spec-correct test expectations)
- **Test count:** 1015 total (up from 939), 1013 passing

### Issue #150 Test Modernization (PR #151)
- **Scope:** Modernized test suite with parameterized tests and parallelization across 6 files
- **Phase 1 — NullValidationTests:** 18 null/empty pairs converted to `[Theory]`, class migrated to `PptxTestBase`
- **Phase 2 — BoundaryConditionTests:** 6 boundary-condition groups parameterized, class migrated to `PptxTestBase`
- **Phase 3 — Parallelization:** Added `xunit.runner.json` with aggressive parallel mode; test time halved (~15s → ~7s)
- **Phase 4 — Additional Theories:** ValidationTests (3→1), TextFormattingTests (9→4), CompareTests (3→1)
- **Net result:** -342 lines, 1022/1022 passing, zero behavioral changes
- **Key patterns:**
  - `string? filePath` parameter with `[InlineData(null)]`/`[InlineData("")]` for null/empty file path pairs
  - `bool useNull` flag for array-typed parameters (mutations, headers, updates)
  - `PptxTestBase` provides `Service`, `CreateMinimalPptx()`, `CreatePptxWithSlides()`, `TrackTempFile()`
  - xUnit v3 supports `[InlineData(null)]` for nullable string parameters