# Squad Decisions

## Active Decisions

### Copilot PR Review Directive (2026-03-16)
**By:** Jon (user directive)
**Status:** Active

All PRs created by squad members must request review from Copilot. Use `--reviewer copilot` on `gh pr create` calls.

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
