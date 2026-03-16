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
