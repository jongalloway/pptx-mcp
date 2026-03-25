---
name: "openxml-test-fixtures"
description: "Build realistic PPTX fixtures for service and tool tests"
domain: "testing"
confidence: "high"
source: "issue-6"
---

## Context

`tests/PptxTools.Tests/TestPptxHelper.cs` is the shared test-fixture builder for pptx-tools. Use it whenever a test needs a valid presentation with realistic combinations of titles, body placeholders, and images.

## Pattern

- Use `TestPptxHelper.CreateMinimalPresentation(...)` for single-slide happy-path tests.
- Use `TestPptxHelper.CreatePresentation(path, slides)` with `TestSlideDefinition` for richer fixtures.
- Model text content with `TestTextShapeDefinition`, using `PlaceholderType = PlaceholderValues.Body` when you want bullet-like talking-point behavior.
- Set `IncludeImage = true` to validate image-only slides without hand-building picture relationships.
- Keep tool tests thin and let service tests cover behavior details.

## Example

```csharp
TestPptxHelper.CreatePresentation(path,
[
    new TestSlideDefinition
    {
        TitleText = "Launch Plan",
        TextShapes =
        [
            new TestTextShapeDefinition
            {
                PlaceholderType = PlaceholderValues.Body,
                Paragraphs =
                [
                    "Ship the MCP tool",
                    "Validate output"
                ]
            }
        ]
    },
    new TestSlideDefinition
    {
        IncludeImage = true
    }
]);
```

## Anti-Patterns

- Do not hand-build OpenXML slide trees inside individual tests when the helper can express the fixture.
- Do not use invalid placeholder or image setup; start from the helper's valid presentation scaffold.
