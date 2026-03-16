# Squad Decisions

## Active Decisions

### Phase 2 Decomposition (2026-03-16)
**Lead:** McCauley  
**Status:** Approved

Phase 2 ("Content Writing & Intelligent Updates") has been decomposed into **5 GitHub issues** (#15–#19), assigned to the "Phase 2 — Content Writing & Updates" milestone.

| # | Title | Owner | Type | Goal | Depends On |
|---|-------|-------|------|------|-----------|
| #19 | Implement pptx_update_slide_data tool | cheritto | Tool | 2A | Phase 1 core |
| #17 | Test pptx_update_slide_data with real metric slides | shiherlis | Testing | 2A | #19 |
| #18 | Design multi-source composition example | copilot | Docs/Example | 2B | None (parallel) |
| #15 | E2E test: multi-source update scenario | shiherlis | Testing | 2B | #19, #18 |
| #16 | Update documentation: Phase 2 tools and workflows | copilot | Docs | 2A+2B | #17, #15 |

**Dependency Chain:**
```
#19 (tool) → #17 (tool testing) ─┐
                                   ├→ #15 (E2E test)
#18 (composition example) ────────┘
                                   │
                         #16 (docs, closes last)
```

**Rationale:** Core tool (#19) unblocks testing and multi-source composition. Testing (#17, #15) validates PowerPoint compatibility and end-to-end scenarios. Documentation and examples (#18, #16) demonstrate Phase 2 capabilities and close the milestone.

## Governance

- All meaningful changes require team consensus
- Document architectural decisions here
- Keep history focused on work, decisions focused on direction
