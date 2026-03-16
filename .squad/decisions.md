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

---

### Phase 1 Documentation Issues (2026-03-17)

**Lead:** McCauley  
**Status:** Active

#### Summary
Created 5 user-facing documentation issues to support Phase 1 launch. These address Jon's request to establish a problem-centric README plus essential docs for end users integrating pptx-mcp into their MCP clients. All assigned to @copilot; no hard dependencies between issues.

#### Issues Created

All assigned to @copilot (Coding Agent), Phase 1 milestone, labels: `documentation`, `phase-1`, `squad`, `squad:copilot`:

- **#10 — Rewrite README:** Problem statement + quick install + capabilities + use cases
- **#11 — TOOL_REFERENCE.md:** Complete alphabetical tool reference with parameters, returns, and JSON examples
- **#12 — QUICKSTART.md:** Zero-to-working guide (prerequisites, install, MCP config, first command, troubleshooting)
- **#13 — CLIENT_SETUP.md:** Step-by-step client configuration (Claude Desktop, VS Code, CLI, local LLMs)
- **#14 — EXAMPLES.md:** 2–4 real use case walkthroughs with agent prompts and tool workflows

#### Squad Assignment Rationale
- **@copilot** → All 5 documentation issues
- **Rationale:** Copilot charter: small features with clear specs, autonomous work
- **Parallel execution:** No blocking dependencies; all can proceed simultaneously
- **Allows Cheritto & Shiherlis** to work on #6, #7, #8 without waiting for docs

#### Scope Decisions
1. **User-facing only** — No internal architecture, dev guides, or decision docs
2. **Shareable with external teams** — Docs should be linkable to outsiders
3. **Future-proofed** — Installation section notes future NuGet publishing
4. **Runnable examples** — All JSON examples testable, not aspirational

#### Dependencies & Ordering
No hard dependencies; soft ordering for coherence:
1. README (#10) — foundation; can deepen as Phase 1 tools land
2. QUICKSTART (#12) — parallel with README
3. CLIENT_SETUP (#13) — parallel; setup guidance
4. TOOL_REFERENCE (#11) — parallel; reference all 7 current tools (will expand)
5. EXAMPLES (#14) — parallel; grows richer with Phase 1 tools

#### Future Updates
- **When Phase 1 tools complete (#6, #7):** Add to TOOL_REFERENCE and EXAMPLES, highlight in README
- **When NuGet publishing set up:** Update QUICKSTART install section
- **When new clients tested:** Expand CLIENT_SETUP
### Copilot PR Review Directive (2026-03-16)
**By:** Jon (user directive)
**Status:** Active

All PRs created by squad members must request review from Copilot. Use `"@copilot"` (double-quoted) for reviewer assignment:
- On create: `gh pr create ... --reviewer "@copilot"`
- On existing: `gh pr edit {number} --add-reviewer "@copilot"`
Note: `--reviewer copilot` (unquoted, no @) fails with "not found".

### Ralph: Copilot Branch → PR Detection (2026-03-16)
**By:** Jon (user directive)
**Status:** Active

During Ralph's work-check cycle (Step 1 scan), Ralph must check for `copilot/*` remote branches that do **not** have a corresponding open PR. GitHub's @copilot coding agent may complete work and push to `copilot/*` branches without creating PRs.

**Detection:**
```bash
# Fetch latest remote state
git fetch --all --prune

# List copilot/* remote branches
git branch -r --list 'origin/copilot/*'

# Cross-reference against open PRs
gh pr list --state open --json headRefName --jq '.[].headRefName'
```

Any `copilot/*` branch not in the open PR list → create a PR:
- Match the branch to an issue (check branch name slug against issue titles/numbers)
- `gh pr create --base main --head {branch} --title "{issue title}" --body "Closes #{issue_number}"`
- Log the PR creation in Ralph's status report

**Priority:** Run this check in every Ralph cycle alongside the existing issue/PR scans.

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
