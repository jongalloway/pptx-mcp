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

### Issue #138: Documentation Complete (2026-04-15)
- **Status:** ✅ Complete and shipped
- **Deliverables:** 4 new guides + README update
  - EMU_GUIDE.md (5.5 KB) — English Metric Unit reference, precision, tool coverage
  - SHAPE_RESOLUTION_GUIDE.md (11.2 KB) — Target shapes by name/index/placeholder, discovery methods, scenarios
  - TROUBLESHOOTING.md (13.2 KB) — 10 category sections (installation, files, format, shapes, text, charts, layouts, performance, testing, CI/CD), quick reference
  - CONTRIBUTING.md (18.7 KB) — Development guide with 5-step "add new tool" example, project structure, test patterns, CI/CD, code style
  - README.md updated (+7 lines) — Added "Guides" section linking to new docs, updated "Contributing / Dev Setup" section
- **Quality:** All 1,227 tests pass, zero regressions, examples verified copy-pasteable, internal links checked, voice consistent across guides
- **McCauley's companion finding:** CLI shipped but undocumented; #138 fills 4 knowledge gaps but README still needs explicit CLI quick-start section as P0 priority
- **Decision:** Merged `.squad/decisions/inbox/shiherlis-docs-138.md` into `.squad/decisions.md`
- **Orchestration log:** `.squad/orchestration-log/2026-04-15T15-11-07Z-shiherlis.md`
- **Testing scope:** Issue #17 (tool testing) + Issue #15 (E2E scenario)
- **Test cases:** 7 integration tests (edge cases, format preservation, Unicode)
- **E2E scenario:** 4-slide KPI dashboard, dual targeting (shapeName + placeholderIndex), format fidelity verification
- **Quality:** Realistic fixtures (TestPptxHelper), OpenXML Validator zero errors, PowerPoint round-trip verified
- **Coverage:** 66/66 tests passing (up from 52), includes speaker notes integrity check
- **Dependency satisfaction:** Both issues unblocked by #19 (Cheritto's tool) and #18 (Copilot's docs)
- **Result:** Phase 2 testing complete, validates PowerPoint compatibility and multi-source composition pattern

### Issue #138 Documentation Audit & Implementation (2026-04-15)

**Author:** Shiherlis (Tester)  
**Assigned:** Issue #138 (Docs: EMU calculator guide, troubleshooting guide, contributing guide) + #94 context

**Status:** ✅ Complete

**Work Summary:**
- **Audit:** Reviewed existing docs (README, TOOL_REFERENCE, QUICKSTART, PRD, EXAMPLES) and identified gaps in user/contributor guidance
- **Created 4 comprehensive guides** (2,442 total lines):
  1. `docs/EMU_GUIDE.md` (280 lines) — EMU conversion reference with quick lookup tables, standard slide dimensions, formula, positioning scenarios, troubleshooting
  2. `docs/SHAPE_RESOLUTION_GUIDE.md` (560 lines) — Shape targeting methods (by name/index/placeholder), discovery via resources, comparison matrix, real scenarios with code examples
  3. `docs/TROUBLESHOOTING.md` (660 lines) — Comprehensive issue reference: setup/file/format/shape/text/chart/layout/performance/test/CI issues, each with causes and solutions
  4. `docs/CONTRIBUTING.md` (935 lines) — Development guide: setup, project structure, architecture pattern (Models→Services→Tools→MCP), example tool implementation walkthrough, test patterns, CI/CD, code style, PR process
- **Updated:** README.md with guide links, improved Contributing section callout
- **Decision:** Documented in `.squad/decisions/inbox/shiherlis-docs-138.md` for maintainer reference
- **Testing:** All 1,227 tests passing, no regressions
- **Quality:** Commands verified against actual CI workflow (e.g., `dotnet test --solution PptxTools.slnx --configuration Release --no-build`), internal links validated, accuracy cross-checked

**Key decisions made:**
- Used existing docs structure; no unnecessary layout changes
- Test commands match CI exactly (important for contributor documentation accuracy)
- Included runnable code examples in CONTRIBUTING.md (new tool walkthrough, test templates)
- Organized by use case: EMU/Shape for users with tooling questions; TROUBLESHOOTING for error resolution; CONTRIBUTING for developers

**Impact:** Fully resolved issue #138. Lowered barrier to entry for both users (EMU/Shape guides) and contributors (CONTRIBUTING guide). Comprehensive troubleshooting reference reduces friction for support. All test patterns documented for consistency.

**Also addresses:** Issue #94 research scope — test patterns and CI/CD documented in CONTRIBUTING.md for future CLI interface work.