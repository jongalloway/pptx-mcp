# Squad Decisions

## Active Decisions

### Decision: PRD Scope & Phase 1 Prioritization

**Date:** 2026-03-15  
**Owner:** McCauley (Lead)  
**Status:** Active

#### Decision

Create PRD at `docs/PRD.md` with explicit phasing:

- **Phase 1 (Current):** Content reading & extraction
  - Extract talking points from slides
  - Export presentations to markdown
  - Timeline: 2–3 weeks
  - Success criteria: Both tools working + tested on real presentations

- **Phase 2 (Deferred):** Intelligent updates & multi-source composition
  - Data-driven slide updates
  - Multi-MCP orchestration (pptx-mcp + external sources)
  - Timeline: 3–4 weeks (after Phase 1 validation)

#### Rationale

1. **Focus** — Phasing prevents scope creep. Phase 1 delivers measurable value (reading) before attempting complex writes.
2. **Risk** — Writing/updating is harder than reading; Phase 1 validates tool reliability before multi-MCP composition.
3. **Validation** — Phase 1 completion allows Jon to validate the model before Phase 2 investment.
4. **Non-Goals** — Explicitly exclude GUI, legacy Office formats, advanced design features, and multi-document transactions to keep complexity bounded.

#### Recommended Actions

- Create 4 GitHub issues from PRD section 8 (Recommended Issues for Phase 1)
- Begin with tool implementations in parallel; target E2E test by end of week 2
- Schedule Phase 1 validation review with Jon before Phase 2 kickoff

#### Team Input Needed

- Jon: Confirm Phase 1 priorities match his "top bullet points" and "markdown export" goals
- Team: Validate 4-week Phase 2 estimate after Phase 1 completes

## Governance

- All meaningful changes require team consensus
- Document architectural decisions here
- Keep history focused on work, decisions focused on direction
