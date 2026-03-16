# Talking points extraction heuristics

- **Date:** 2026-03-17
- **Author:** Cheritto
- **Related issue:** #6

## Summary
`pptx_extract_talking_points` ranks visible slide text by placeholder hierarchy and bullet-like structure while filtering noise such as presenter-note labels and formatting-only text.

## Why it matters
This keeps the tool aligned with Phase 1's visible-content scope and avoids treating visual-only slides as if they contained meaningful bullets.

## Implementation notes
- Prefer body and object placeholders over title placeholders when ranking.
- Use title text as a fallback only when a slide has no stronger text candidates and no other visual-only content.
- Shared fixture coverage lives in `tests/PptxMcp.Tests/TestPptxHelper.cs` and `tests/PptxMcp.Tests/Services/PresentationServiceTests.cs`.
