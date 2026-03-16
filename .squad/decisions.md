# Squad Decisions

## Active Decisions

### Phase 1 Issue Structure & Squad Assignments (2026-03-16)

**Lead:** McCauley  
**Status:** Active

#### Summary
Created GitHub issue structure for Phase 1 ("Content Reading & Extraction") PRD. Four issues span tool implementation, integration testing, and documentation. Assigned to squad members based on capability fit and charter alignment.

#### Issues Created
- **#6 — Implement pptx_extract_talking_points tool** (Medium complexity, Cheritto)
  - New MCP tool to extract key bullet points from slides
  - Labels: `tool`, `phase-1`, `squad:cheritto`
  - Milestone: Phase 1 — Content Reading & Extraction
  - Acceptance: Structured output, unit tests (5+ cases), integration tested
  
- **#7 — Implement pptx_export_markdown tool** (Medium complexity, Cheritto)
  - New MCP tool to export full presentation to clean markdown
  - Labels: `tool`, `phase-1`, `squad:cheritto`
  - Milestone: Phase 1 — Content Reading & Extraction
  - Acceptance: Valid markdown, unit tests (5+ cases), integration tested
  
- **#8 — E2E test: read real presentation and export markdown** (Low complexity, Shiherlis)
  - Integration test on 3+ real presentations
  - Labels: `testing`, `phase-1`, `squad:shiherlis`
  - Milestone: Phase 1 — Content Reading & Extraction
  - Dependencies: #6, #7
  - Acceptance: 3+ diverse presentations tested, accuracy verified, CI passing
  
- **#9 — Document Phase 1 tools and example workflows** (Low complexity, @copilot)
  - README updates, agent prompts, JSON examples, use case narrative
  - Labels: `documentation`, `phase-1`, `squad:copilot`
  - Milestone: Phase 1 — Content Reading & Extraction
  - Dependencies: #6, #7
  - Acceptance: 2+ example prompts, JSON request/response shown

#### Squad Assignments Rationale
- **Cheritto** → Tool implementations (#6, #7) — Backend Dev charter
- **Shiherlis** → Integration testing (#8) — Tester charter  
- **@copilot** → Documentation (#9) — Small features with specs (auto-assigned, requires review)

#### Milestones & Labels
**Milestones:**
- Phase 1 — Content Reading & Extraction (milestone #1)
- Phase 2 — Content Writing & Updates (milestone #2, created for future use)

**Labels:**
- `tool` — New MCP tool implementations
- `testing` — Test coverage and E2E testing
- `documentation` — Documentation updates
- `phase-1`, `phase-2` — Phase markers
- `squad` — Squad triage marker
- `squad:cheritto`, `squad:shiherlis`, `squad:copilot` — Team assignment labels

#### Scope Decisions
1. **No speaker notes for Phase 1** — Assume visible content only; parking lot for later
2. **Three-issue dependency chain** — Structured to avoid bottlenecks
3. **Minimum test coverage** — 5+ unit tests per tool; E2E covers 3+ presentations
4. **Real-world validation** — Both tools must be tested on real presentations before acceptance

---

### User Directive: Copilot PR Review (2026-03-16)

**By:** Jon Galloway (via Copilot)  
**Status:** Active

**Directive:** Always request a review from Copilot on each PR the squad creates. Use `--reviewer @copilot` (or equivalent) when opening PRs via `gh pr create`.

**Rationale:** User request — captured for team memory

## Governance

- All meaningful changes require team consensus
- Document architectural decisions here
- Keep history focused on work, decisions focused on direction
