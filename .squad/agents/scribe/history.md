# Project Context

- **Owner:** Jon Galloway
- **Project:** pptx-mcp — .NET 10 MCP server for PowerPoint manipulation via OpenXML SDK
- **Stack:** .NET 10, C#, ModelContextProtocol v1.1.0, DocumentFormat.OpenXml v3.3.0, xUnit v3 (MTP)
- **Team:** McCauley (Lead), Cheritto (Backend Dev), Shiherlis (Tester), Nate (Consulting Dev)
- **Created:** 2026-03-16

## Core Context

Agent Scribe initialized and ready for work.

## Recent Updates

📌 **2026-04-10 (2346Z):** Notable state review & config fix
- Processed McCauley's recommendations: config.json teamRoot mismatch + now.md refresh
- Orchestration log created: `.squad/orchestration-log/2026-04-10T23-46-16Z-mccauley-notable-state-review.md`
- Session log created: `.squad/log/2026-04-10T23-46-16Z-notable-state-recommendations.md`
- Decision inbox merged: `mccauley-recommend-notable-state.md` merged into decisions.md, deleted from inbox
- Config fix: `.squad/config.json` teamRoot corrected from D:\Users\Jon\... to C:\Users\Jon\... (critical blocker resolved)
- Cross-agent updates: Scribe history (this), McCauley history (completion)
- Status: Config path resolution functional, next agent spawn will have full context

📌 **2026-03-18 (0930Z):** Tool consolidation workflow completed
- Orchestration logs created (McCauley design, Cheritto impl, McCauley review)
- Session log written
- Decision inbox merged into decisions.md (4 files consolidated, duplicates removed)
- Git commit: fde7c1a (squad decisions updated)
- Status: PR #76 merged, Issue #69 closed, 377 tests passing

📌 Team initialized on 2026-03-16

## Learnings

- Tool consolidation via enum-based action dispatch is clean and SDK-native
- Decision merging centralizes team memory without duplication
- Orchestration logs provide transparent audit trail of agent work
- Squad log provides executive summary for each workflow
