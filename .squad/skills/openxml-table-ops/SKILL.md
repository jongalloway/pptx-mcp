# Skill: OpenXML Table Operations

## When to Use
Creating or modifying DrawingML tables in PowerPoint files via OpenXML SDK.

## Table Creation Pattern

Tables in PPTX are `GraphicFrame` elements, NOT `Shape` elements.

Structure: `P.GraphicFrame > A.Graphic > A.GraphicData > A.Table > A.TableRow > A.TableCell`

### Critical Rules
1. **GraphicData URI** must be exactly: `http://schemas.openxmlformats.org/drawingml/2006/table`
2. **TableGrid column count** must equal cell count in every row — pad short rows
3. **Every TableCell** must have: `TextBody(BodyProperties(), ListStyle(), Paragraph(Run(Text(...)), EndParagraphRunProperties()))` + `TableCellProperties()`
4. **Shape IDs** must be unique — use `GetMaxShapeId(shapeTree) + 1`
5. **Append** GraphicFrame to ShapeTree (same as Shape/Picture)

### Name Lookup
Tables use `frame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name` — NOT `NonVisualShapeProperties`.

### Cell Text Update
- Preserve `TableCellProperties` (clone before replacing)
- Replace `TextBody` entirely with fresh structure (BodyProperties + ListStyle + Paragraph)
- Do NOT try to reuse existing `TextBody` — rebuild is safer

## Reference Files
- `src/PptxMcp/Services/PresentationService.cs` — `InsertTable()`, `UpdateTable()`, `BuildTableRow()`
- `tests/PptxMcp.Tests/TestPptxHelper.cs` — `CreateTable()` (test fixture builder)
