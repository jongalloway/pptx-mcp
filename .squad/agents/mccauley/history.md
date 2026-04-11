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

### Notable State Review Process (2026-04-10)
- McCauley identified critical config.json path mismatch during squad state audit
- Scribe processed recommendations: wrote orchestration log, session log, merged decisions, updated agent history
- config.json fix unblocks next agent spawn and restores coordinator path resolution
- Squad state review workflow: audit → document → merge → cross-update → commit

## Recent Updates

📌 **2026-04-11 (0000Z):** Squad state cleanup final pass completed
- McCauley verified config.json (still good), refreshed now.md (focus_area + active_issues updated)
- Audit: No additional tightly-coupled stale metadata found
- Orchestration log: `.squad/orchestration-log/2026-04-11-000010-mccauley-squad-state-cleanup.md`
- Session log: `.squad/log/2026-04-11-000010-squad-state-cleanup.md`
- Decisions merged: `mccauley-squad-state-cleanup.md` integrated into decisions.md
- Status: ✅ Complete, squad metadata now current and accurate

📌 **2026-04-10 (2346Z):** Notable state review completed
- McCauley's recommendations processed: config.json fix + now.md refresh tracked
- Orchestration log: `.squad/orchestration-log/2026-04-10T23-46-16Z-mccauley-notable-state-review.md`
- Session log: `.squad/log/2026-04-10T23-46-16Z-notable-state-recommendations.md`
- Decisions merged: `mccauley-recommend-notable-state.md` integrated into decisions.md
- Config fix: teamRoot corrected (D: → C:), path resolution functional
- Status: ✅ Complete, ready for next agent spawn

### Phase 4 Wave Decomposition & Sequencing (2026-03-24)
- **3-Tier Wave Strategy:** Wave 1 (Core analysis: #80, #81, #82), Wave 2 (Enhancement & consolidation), Wave 3 (Polish & documentation)
- **Squad Reassignments:** Cheritto → 3-tool implementation sprint (wave 1), Shiherlis → 92-test parallel validation, Nate → OpenXML research/skill capture, Coordinator → PR review & docs
- **Risk Assessment:** Identified PAT permissions bottleneck for automated PR creation/label updates; contingency: manual intervention workflow
- **Execution:** Wave 1 complete with all deliverables on branch/PR (0 blockers on implementation, all items ready for integration)
- **Pattern Success:** Wave-based sequencing with parallel implementation/testing streams scales well for tooling work; allows PAT issues to surface early without blocking squad

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

### Squad Configuration Review (2026-03-27)
- **Item 1 (now.md):** Stale since bootstrap (2026-03-16). Generic "Initial setup" still in place; empty active_issues. **Low risk** — informational only, doesn't affect coordinator. Recommend refresh at next planning cycle to surface current focus (feature waves #115, #116, #121, #125) and improve squad orientation.
- **Item 2 (config.json):** CRITICAL mismatch — `teamRoot` points to `D:\Users\Jon\Documents\GitHub\pptx-tools` but repo is at `C:\Users\Jon\...`. Coordinator uses this path to resolve `.squad/*` references. **Active blocker** — agents spawning will fail soft on charter/decisions resolution if not fixed. Fix immediately before next squad work.
- **Coordinator impact:** config.json mismatch causes path resolution failures; agents degrade to reduced context (missing charter, team routing broken).

### Squad State Cleanup — Final Pass (2026-04-11)
- **config.json:** ✅ Verified still correct (teamRoot = `C:\Users\Jon\Documents\GitHub\pptx-tools`). No action needed.
- **now.md:** ✅ Refreshed with current project state — Updated focus_area to "Advanced feature waves & test modernization (waves #125–#150)" and active_issues to [125, 135, 140, 150]. Timestamp updated to 2026-04-11.
- **Additional cleanup audit:** Checked `.squad/identity/wisdom.md` (empty template—intended), `.squad/decisions/inbox/` (clean, previously merged), `.squad/casting-*.json` files (inert config—no cleanup needed), agent histories (all recently updated), `.squad/routing.md` and `.squad/team.md` (current). No additional low-risk cleanup items identified.
- **Outcome:** Squad metadata now current and accurate. Coordinator and agent spawning fully functional.

