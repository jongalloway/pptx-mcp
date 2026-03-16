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

<!-- Append new learnings below. Each entry is something lasting about the project. -->
