# Project Context

- **Owner:** Jon Galloway
- **Project:** pptx-mcp — .NET 10 MCP server for PowerPoint manipulation via OpenXML SDK
- **Stack:** .NET 10, C#, ModelContextProtocol v1.1.0, DocumentFormat.OpenXml v3.3.0, xUnit v3 (MTP), Microsoft.Extensions.Hosting v10.0.5
- **Architecture:** Console app with stdio transport. Models → Services (PresentationService) → Tools (PptxTools) → MCP server
- **Key files:** src/PptxMcp/Tools/PptxTools.cs (169 lines, 7 tools), src/PptxMcp/Services/PresentationService.cs (464 lines, all OpenXML ops)
- **Build:** `dotnet build PptxMcp.slnx --configuration Release`
- **Test:** `dotnet test --solution PptxMcp.slnx --configuration Release --no-build`
- **Reference repos:** jongalloway/dotnet-mcp (MCP patterns), jongalloway/MarpToPptx (OpenXML patterns)
- **Created:** 2026-03-16

## Learnings

### Phase 1 Issue Creation (2026-03-16)
- Assigned #6 & #7 (pptx_extract_talking_points, pptx_export_markdown) for implementation
- Both are Medium complexity, can be parallelized
- Tool implementations must be integration-tested on real presentations before acceptance
- Depends on Shiherlis for E2E validation (#8) and @copilot for documentation (#9)
- All issues reference docs/PRD.md for success criteria alignment

<!-- Append new learnings below. Each entry is something lasting about the project. -->

### Markdown Export Tool (2026-03-16)
- `pptx_export_markdown` should keep tool logic thin and delegate markdown generation to `PresentationService.ExportMarkdown(...)`.
- Phase 1 markdown export excludes speaker notes even though notes are available elsewhere in `PresentationService`.
- Exported images belong in a sibling `<markdown-file>_images` folder and markdown should reference them with relative forward-slash paths for portability.
- Realistic PPTX fixtures need explicit paragraph/table/image construction in `TestPptxHelper` to validate nested bullets, tables, and embedded media.
