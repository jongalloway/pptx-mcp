# Project Context

- **Owner:** Jon Galloway
- **Project:** pptx-mcp — .NET 10 MCP server for PowerPoint manipulation via OpenXML SDK
- **Stack:** .NET 10, C#, ModelContextProtocol v1.1.0, DocumentFormat.OpenXml v3.3.0, xUnit v3 (MTP), Microsoft.Extensions.Hosting v10.0.5
- **Test project:** tests/PptxMcp.Tests/ — xUnit v3 on Microsoft Testing Platform
- **Test command:** `dotnet test --solution PptxMcp.slnx --configuration Release --no-build` (uses `--filter-method` not `--filter`)
- **Reference repos:** jongalloway/dotnet-mcp (test patterns), jongalloway/MarpToPptx (OpenXML test fixtures)
- **Created:** 2026-03-16

## Learnings

### Phase 2 Assignments (2026-03-16)
- **Issue #17 (shiherlis assigned):** Test pptx_update_slide_data with real metric slides — validates PowerPoint compatibility and edge cases for Goal 2A
- **Issue #15 (shiherlis assigned):** E2E test multi-source update scenario — validates full composition workflow (Goal 2B)
- Dependency: Both #17 and #15 depend on #19 (core tool implementation by cheritto) being complete
- Timeline: Phase 2 estimated 3–4 weeks after Phase 1 stabilization
- Test approach: Use TestPptxHelper.cs fixtures for realistic metric slides and multi-source composition patterns

<!-- Append new learnings below. Each entry is something lasting about the project. -->
