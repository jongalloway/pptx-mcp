# Project Context

- **Owner:** Jon Galloway
- **Project:** pptx-mcp — .NET 10 MCP server for PowerPoint manipulation via OpenXML SDK
- **Stack:** .NET 10, C#, ModelContextProtocol v1.1.0, DocumentFormat.OpenXml v3.3.0, xUnit v3 (MTP), Microsoft.Extensions.Hosting v10.0.5
- **Architecture:** Console app with stdio transport. Models → Services (PresentationService) → Tools (PptxTools) → MCP server
- **7 MCP tools:** pptx_list_slides, pptx_list_layouts, pptx_add_slide, pptx_update_text, pptx_insert_image, pptx_get_slide_xml, pptx_get_slide_content
- **Build:** `dotnet build PptxMcp.slnx --configuration Release`
- **Test:** `dotnet test --solution PptxMcp.slnx --configuration Release --no-build` (MTP runner, uses `--filter-method` not `--filter`)
- **Reference repos:** jongalloway/dotnet-mcp (MCP patterns), jongalloway/MarpToPptx (OpenXML patterns)
- **Created:** 2026-03-16

## Learnings

### PRD Structure & Scope (2026-03-15)
- Created PRD at `docs/PRD.md` based on PR #1 bootstrap and Jon's vision
- Phase 1 (Content Reading) focuses on two high-value tools: extract talking points + export markdown
- Phase 2 (Intelligent Updates) deferred pending Phase 1 validation; planned for multi-source composition (pptx-mcp + external data MCPs)
- **Key decision:** Non-goals explicitly exclude GUI, legacy formats, and advanced design features to keep scope bounded
- **Recommended 4 GitHub issues** for Phase 1: two tool implementations, one E2E test, one docs pass
- Timeline estimate: 2–3 weeks Phase 1, 3–4 weeks Phase 2 (estimate includes +20% buffer)

### Phase 1 Issue Creation & Triage (2026-03-16)
- Created two GitHub milestones: Phase 1 (milestone #1) and Phase 2 (milestone #2)
- Created 9 labels for issue organization: `tool`, `testing`, `documentation`, `phase-1`, `phase-2`, `squad`, and squad role labels
- Created 4 GitHub issues for Phase 1 work:
  - **#6 (Cheritto):** pptx_extract_talking_points implementation
  - **#7 (Cheritto):** pptx_export_markdown implementation
  - **#8 (Shiherlis):** E2E testing with real presentations
  - **#9 (@copilot):** Phase 1 docs + example workflows
- Assignment logic: Tool work → Cheritto (backend dev charter), testing → Shiherlis (tester charter), docs → @copilot (small features with specs, auto-assigned, requires review)
- Dependency chain: #6 & #7 independent, #8 depends on both, #9 depends on both
- Decision document written to `.squad/decisions/inbox/mccauley-prd-phase1-issues.md`
- All issues reference PRD success criteria and use acceptance checklists for clarity
