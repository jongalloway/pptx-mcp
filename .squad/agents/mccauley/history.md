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

### Phase 1 Documentation Strategy (2026-03-17)
- Jon requested README rewrite and user-facing docs creation
- Created 5 new documentation issues, all Phase 1 milestone, assigned to @copilot:
  - **#10:** README rewrite — problem-centric structure (WHY + HOW + capabilities + use cases)
  - **#11:** TOOL_REFERENCE.md — comprehensive tool reference with JSON examples
  - **#12:** QUICKSTART.md — zero-to-working guide for new users
  - **#13:** CLIENT_SETUP.md — MCP client configuration (Claude Desktop, VS Code, CLI, local LLMs)
  - **#14:** EXAMPLES.md — real use case walkthroughs with agent prompts and tool workflows
- **Scope clarity:** Documentation is user-facing only (no internal architecture or decision docs)
- **Key decision:** Installation section notes future NuGet publishing; docs are future-proof
- **No hard dependencies** between docs; can be done in parallel. Will grow richer as Phase 1 tools complete
- Decision document written to `.squad/decisions/inbox/mccauley-phase1-documentation-issues.md`
### Phase 2 Decomposition (2026-03-16)
- Decomposed Phase 2 ("Content Writing & Intelligent Updates") into 5 GitHub issues (#15–#19)
- **Core tool work:** Issue #19 (cheritto) implements `pptx_update_slide_data` tool for data-driven slide updates (Goal 2A)
- **Testing:** Issue #17 (shiherlis) validates pptx_update_slide_data with real metric slides; Issue #15 (shiherlis) validates multi-source E2E scenario (Goal 2B)
- **Docs & examples:** Issue #18 (copilot) designs composition pattern demo; Issue #16 (copilot) updates all documentation
- **Dependency chain:** #19 → #17 → #15 + #18 → #16 (docs closes last)
- All issues routed to appropriate team members; all assigned to "Phase 2 — Content Writing & Updates" milestone
- Non-duplication verified: Phase 1 issues #6–#14 are independent; no overlap with Phase 2

### Phase 3 Planning & Feature Ranking (2026-03-17)

- Led Phase 3 planning session with Nate's input on prior-art research
- Ranked 7 features by product impact, complexity, and sequencing: batch refresh → template-aware authoring → tables → picture placeholders → notes → chart refresh → slide organization
- Consulted with Nate (per Jon directive) on prior-art research from MarpToPptx and dotnet-mcp; aligned on feasibility and implementation patterns
- Created 7 GitHub issues (#34–#40) under Phase 3 milestone with comprehensive ownership, acceptance criteria, and dependencies
- Recorded decision to continue McCauley+Nate partnership for major architectural decisions (worked well for aligned thinking, caught gotchas)
- Scope cuts: no full theme/master editing, no net-new chart authoring (refresh existing only).

### Quality & Housekeeping Phase (2026-03-18)

**Status:** Proposed (blocking gate before Phase 3 ramps up)

- Analyzed post-Phase-2 codebase: 260 tests passing, 68 build warnings, ~3,100 core LOC
- Identified three high-priority quality issues:
  1. **Code duplication:** 27× identical file-check + try-catch pattern in PptxTools.cs; 22× inline JsonSerializerOptions; 2× batch failure construction
  2. **Documentation staleness:** README tool list incomplete (table tools from Phase 2 not advertised); TOOL_REFERENCE.md and EXAMPLES.md need sync
  3. **Test coverage gaps:** Missing null/empty input validation, boundary conditions (slide indices, shape IDs, table columns), FileNotFound parameterization
- Assessed OpenXML patterns as consistent and well-established (no action needed); dead code audit found none
- **Created 8 GitHub issues (#48–#56)** under new "Quality & Housekeeping" milestone (tier 1: Q1–Q3 are blocking gates; tier 2: Q4–Q7 are polish; tier 3 deferred)
- Proposed schedule: Week 1 (Q1–Q3 unblock Phase 3), Week 2 (Q4–Q7 polish)
- Team assignments: Cheritto (Q2, Q4 refactoring), Shiherlis (Q3, Q5, Q8 testing), @copilot (Q1, Q6, Q7 docs)
- Decision document written to `.squad/decisions/inbox/mccauley-quality-phase.md`
- **Key finding:** Phase 2 delivered production-ready table operations; quality pass ensures Phase 3 starts with a clean foundation.

### 2026-03-17T06:07Z: Quality & Housekeeping Phase Finalized

- Completed full codebase quality analysis: baseline, duplication patterns, documentation gaps, test coverage assessment
- Outcome: 8 GitHub issues (#48–#56) under "Quality & Housekeeping" milestone
  - Tier 1 (Must Do): Q1–Q3 blocking gates (documentation, boilerplate extraction, validation tests)
  - Tier 2 (Should Do): Q4–Q7 polish (consolidation, boundary tests, doc sync)
  - Tier 3 (Nice to Have): Q8–Q10 deferred (parameterized tests, stress tests, squad archive)
- Decision: Commit to Tier 1 + Tier 2 (9 hours, 1–2 weeks part-time) before Phase 3 ramps up
- Team assignments aligned: Cheritto (refactoring), Shiherlis (testing), @copilot (docs), McCauley (oversight)
- Success criteria: All Tier 1+2 closed, 260+ tests, 68 or fewer warnings, docs reflect all tools, zero Phase 2 regression
- Orchestration log written to `.squad/orchestration-log/2026-03-17T0607Z-mccauley.md`
- Session log written to `.squad/log/2026-03-17T0607Z-quality-phase.md`
- Decisions merged to decisions.md; inbox files deleted
