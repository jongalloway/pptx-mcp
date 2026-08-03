---
name: "openxml-text-updates"
description: "Preserve PowerPoint formatting when replacing text in existing shapes"
domain: "openxml-updates"
confidence: "high"
source: "issue-19 implementation"
---

## Context

Use this pattern when a pptx-tools feature needs to replace text in an existing PowerPoint shape without breaking formatting or package structure.

## Pattern

- Resolve the target shape first from `pptx_get_slide_content` metadata, preferring the shape name (`p:cNvPr/@name`) and using a deterministic index fallback only when needed.
- Work with existing `p:sp` elements instead of recreating the slide shape.
- Clone the existing `TextBody`'s `BodyProperties` and `ListStyle` before replacing content.
- Rebuild paragraphs from the existing paragraph templates so paragraph properties, bullet levels, and run formatting survive the update.
- When a shape has no `TextBody`, insert a new one in the same structural slot PowerPoint expects: after `spPr`/`style` and before `extLst`.
- Save only the modified slide part after the replacement.
- For deck-wide refreshes, open the `PresentationDocument` once, reuse the same slide-id snapshot for every mutation, and save each touched slide only once after the batch finishes.

## Example

```csharp
var existingTextBody = shape.TextBody ?? new TextBody(
    new A.BodyProperties(),
    new A.ListStyle(),
    new A.Paragraph(new A.EndParagraphRunProperties()));

var replacementTextBody = new TextBody(
    (A.BodyProperties)existingTextBody.BodyProperties.CloneNode(true),
    (A.ListStyle)existingTextBody.ListStyle.CloneNode(true));
```

## Anti-Patterns

- Do not wipe the shape and append a bare `A.Run(new A.Text(...))` unless you are willing to lose paragraph and run formatting.
- Do not recreate slide parts or relationships for a text-only update.
- Do not guess at a target shape when multiple names match; return a recoverable failure instead.
- Do not reopen and resave the `.pptx` once per mutation when a workflow can batch several text updates together.
