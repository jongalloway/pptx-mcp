# IMPLEMENTATION MAP: pptx_insert_table and pptx_update_table

## EXECUTIVE SUMMARY
This codebase uses OpenXML SDK with a clean 3-layer architecture:
- **Tools Layer** (PptxTools.cs): MCP-compliant wrappers; JSON return; no exceptions thrown
- **Service Layer** (PresentationService.cs): Business logic; handles document I/O; all errors caught
- **Models** (Models/*.cs): DTOs for structured results and requests

Tables in PPTX are NOT Shape elements—they are GraphicFrame elements wrapping A.Table (OpenXML Drawing namespace). This is crucial.

---

## KEY FILES & PATTERNS

### 1. TOOL PATTERN (PptxTools.cs, 418 lines)
**Location**: C:\Users\Jon\Documents\GitHub\pptx-tools\src\PptxTools\Tools\PptxTools.cs

**Attributes & XML Docs**:
- Attribute: [McpServerTool(Title = "Display Name", ReadOnly = true/false, Idempotent = true/false)]
- XML docs: one-liner summary + <param> for each parameter
- Return: always Task<string> (JSON for complex results, plain error text for failures)

**Example** (pptx_update_slide_data, lines 94-150):
- XML doc describes all parameters including EMU units
- File check at tool level
- Try-catch wraps service call
- Returns JsonSerializer.Serialize(..., new JsonSerializerOptions { WriteIndented = true })

**Pattern for Table Tools**:
- Use 1-based slideNumber (not slideIndex) for consistency
- Return JSON result DTO, not plain text
- Check file before calling service
- Let service raise exceptions; catch and return as error string

---

### 2. SERVICE LAYER PATTERNS (PresentationService.cs, ~1200 lines)
**Location**: C:\Users\Jon\Documents\GitHub\pptx-tools\src\PptxTools\Services\PresentationService.cs

**Write Operation Lifecycle** (see AddSlide, InsertImage, UpdateSlideData):
1. Open: PresentationDocument.Open(filePath, true) [editable mode]
2. Locate: GetSlidePart(doc, slideIndex) or GetSlidePart(doc, slideIds, slideNumber-1)
3. Find/Modify: Locate target, apply changes
4. Save: slide.Save() (or batch: collect parts, save once each)
5. Return: DTO with Success, previous/new values, message

**Slide Lookup Helpers**:
- GetSlidePart(PresentationDocument, int slideIndex) — line 424
- GetSlidePart(PresentationDocument, IReadOnlyList<SlideId>, int slideIndex) — line 430
- GetSlideIds(PresentationDocument) — line 439
- All validate ranges; throw ArgumentOutOfRangeException with count info

**Shape ID Management**:
- GetMaxShapeId(ShapeTree) — line 405
- Scans all children for NonVisualDrawingProperties?.Id?.Value
- Returns max; increment by 1 for new shapes
- Handles Shape, Picture, GraphicFrame, GroupShape, ConnectionShape

**Text Replacement with Formatting Preserved** (lines 614-697):
- ReplaceShapeTextPreservingFormatting(Shape, string)
- Clones BodyProperties, ListStyle, paragraph templates
- Preserves formatting; replaces only text runs
- Used by UpdateSlideData; NOT applicable to tables

**TABLE-SPECIFIC**:
- ExtractGraphicFrame() — line 878: detects if frame contains table
- ExtractTableRows(A.Table) — line 910: parses row/cell text
- Tables nested: GraphicFrame > Graphic > GraphicData > A.Table > A.TableRow > A.TableCell > A.TextBody

---

### 3. RESULT/REQUEST DTOs (Models/)
**Location**: C:\Users\Jon\Documents\GitHub\pptx-tools\src\PptxTools\Models/

**SlideDataUpdateResult** (SlideDataUpdateResult.cs):
`csharp
public record SlideDataUpdateResult(
    bool Success,
    int SlideNumber,                // 1-based
    string? RequestedShapeName,
    int? RequestedPlaceholderIndex,
    string? MatchedBy,              // "shapeName", "placeholderIndex", "placeholderIndexFallback"
    string? ResolvedShapeName,
    int? ResolvedShapeIndex,        // 0-based among text-capable shapes
    uint? ResolvedShapeId,          // OpenXML shape ID (unique in presentation)
    string? PlaceholderType,        // "Title", "Body", etc.
    uint? LayoutPlaceholderIndex,
    string? PreviousText,           // Before modification
    string NewText,                 // What was requested
    string Message);                // Human-readable status
`

**BatchUpdateMutation** (request DTO):
`csharp
public record BatchUpdateMutation(
    int SlideNumber,                // 1-based
    string ShapeName,
    string NewValue);
`

**PATTERN FOR TABLE DTOs** (create new):
- TableInsertResult: Success, SlideNumber, TableName, TableShapeId, TableIndex, RowCount, ColumnCount, Message
- TableUpdateResult: Success, SlideNumber, TableName, MatchedBy, PreviousRowCount, NewRowCount, Message
- Follow SlideDataUpdateResult pattern for consistency

---

### 4. EXISTING TABLE SUPPORT (Already in codebase)

**Reading Tables** (lines 878-921):
`csharp
private static ShapeContent ExtractGraphicFrame(P.GraphicFrame frame)
{
    // ...
    var table = graphicData.GetFirstChild<A.Table>();
    if (table is not null)
        tableRows = ExtractTableRows(table);
    // ShapeType = tableRows is not null ? "Table" : "GraphicFrame"
}

private static IReadOnlyList<IReadOnlyList<string>> ExtractTableRows(A.Table table)
{
    var rows = new List<IReadOnlyList<string>>();
    foreach (var row in table.Elements<A.TableRow>())
    {
        var cells = new List<string>();
        foreach (var cell in row.Elements<A.TableCell>())
            cells.Add(cell.InnerText);
        rows.Add(cells);
    }
    return rows;
}
`

**ShapeContent** already supports tables:
- ShapeType: "Table" | "GraphicFrame" | etc.
- TableRows: IReadOnlyList<IReadOnlyList<string>>? (null for non-table shapes)
- pptx_get_slide_content already returns tables in this format

**Markdown Export** (lines 1068-1089):
- AppendTableMarkdown() writes markdown pipes
- Handles ragged row widths (pads with empty cells)
- Does NOT preserve formatting, only text
- Your updates must not break this

---

### 5. TEST HELPERS & FIXTURES

**TestPptxHelper.cs** (340 lines):
- CreateMinimalPresentation(filePath, titleText)
- CreatePresentation(filePath, TestSlideDefinition[])
- BuildSlide(slidePart, definition) — assembles shapes

**CreateTable() method** (lines 278-324):
Shows complete table creation pattern:
1. Create A.Table with TableProperties { FirstRow=true, BandRow=true }
2. Append A.TableGrid with GridColumn for each column
3. For each row: create A.TableRow { Height = ... }
4. For each cell: A.TableCell(TextBody(...), TableCellProperties())
5. Wrap in P.GraphicFrame with NonVisualGraphicFrameProperties, Transform, Graphic
6. GraphicData Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"

**TestTableDefinition**:
`csharp
public sealed class TestTableDefinition
{
    public string? Name { get; init; }
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; init; } = [];  // Row > cell text
    public long? X, Y, Width, Height { get; init; }  // EMUs
}
`

**Test Organization**:
- UpdateSlideDataTests.cs: shape resolution, formatting preservation, batch updates, AssertPresentationCompatible()
- MarkdownExportTests.cs: includes table export test (lines 129-156)
- PptxPhase1E2eTests.cs: end-to-end with table (lines 73-84)
- PptxToolsTests.cs: tool-level JSON and error handling

---

## CRITICAL POWERPOINT COMPATIBILITY GOTCHAS

### For pptx_insert_table:
1. **TableGrid must match structure**:
   - GridColumn count = max cells in any row
   - Pad short rows with empty cells
   - All rows must have same column count

2. **Unique shape IDs**:
   - Call GetMaxShapeId() to find next available
   - PowerPoint crashes on duplicate IDs

3. **Transform positioning** (X, Y in EMUs):
   - Use 914400 (1 inch) as default X
   - Stack tables with Y offset + height + spacing
   - Height should fit cell text (cells clip if too small)

4. **Cell TextBody structure**:
   - Each cell MUST have: TextBody > BodyProperties > ListStyle
   - Each TextBody MUST have >= 1 Paragraph
   - Paragraph must have Run or EndParagraphRunProperties
   - Missing = XML corruption, unreadable file

5. **GraphicData Uri**:
   - EXACT: "http://schemas.openxmlformats.org/drawingml/2006/table"
   - One character wrong = table invisible

6. **Append to shapeTree**:
   - shapeTree.Append(graphicFrame) adds to end
   - Order doesn't matter for rendering, but affects z-order

### For pptx_update_table:
1. **Tables are GraphicFrame, not Shape**:
   - Name search: frame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name
   - Can't use existing text-shape logic directly

2. **Cell text replacement**:
   - Replace TextBody INSIDE A.TableCell
   - Preserve A.TableCellProperties (borders, fills)
   - Decision: preserve cell formatting or strip it? (Recommend: strip for simplicity in MVP)

3. **Structural changes (add/remove rows/cols)**:
   - **Very hard**. Requires recalc TableGrid, row/cell consistency
   - **Recommend**: MVP supports only cell-text updates, not structure

4. **In-place vs. replacement**:
   - Option A: Modify cells in-place (simpler, but risky if schema mismatch)
   - Option B: Replace entire table (safer, aligns with test helper pattern)
   - For MVP, use Option A: locate table > modify cell TextBodies > save

5. **Shape tree integrity**:
   - If removing/inserting, use ReplaceChild() not Remove() + Append()
   - Maintain z-order and ID uniqueness

---

## FILES NEEDING EDITS

### PRIMARY (Must create/edit):

1. **src/PptxTools/Tools/PptxTools.cs** [existing]
   - Add: pptx_insert_table()
   - Add: pptx_update_table()
   - OR create: PptxTools.Tables.cs (partial class)

2. **src/PptxTools/Services/PresentationService.cs** [existing]
   - Add: InsertTable(filePath, slideIndex, rows, x, y, width, height, name?)
   - Add: UpdateTable(filePath, slideIndex, tableName, tableIndex?, rows?, updateMode?)

3. **src/PptxTools/Models/TableInsertResult.cs** [NEW]

4. **src/PptxTools/Models/TableUpdateResult.cs** [NEW]

5. **tests/PptxTools.Tests/Services/TableOperationTests.cs** [NEW]
   - Test InsertTable: row/col counts, ID uniqueness, positioning
   - Test UpdateTable: cell updates, compatibility
   - Test with real metric deck (UpdateSlideDataTests.cs pattern)

6. **tests/PptxTools.Tests/Tools/TableToolsTests.cs** [NEW]
   - Test pptx_insert_table: JSON format, file errors, parameter validation
   - Test pptx_update_table: shape resolution, error messages

### SECONDARY (Verify, no changes):

7. **src/PptxTools/Services/PresentationService.cs** [existing]:
   - GetSlideContent() — already extracts tables, no change needed
   - Markdown export — already handles tables, no change needed

8. **tests/PptxTools.Tests/TestPptxHelper.cs** [existing]:
   - CreateTable() already exists
   - TestTableDefinition already exists
   - Reuse for test setup

---

## RECOMMENDED IMPLEMENTATION ORDER

1. Create DTOs (TableInsertResult.cs, TableUpdateResult.cs)
2. Implement PresentationService.InsertTable()
3. Implement PresentationService.UpdateTable()
4. Add pptx_insert_table() tool
5. Add pptx_update_table() tool
6. Write service tests (table creation, cell updates, compatibility)
7. Write tool tests (JSON, errors)
8. Test with existing Phase 1 E2E test (verify markdown export still works)

---

## REFERENCE: Similar Operations in Codebase

| Op | Tool | Service | Result DTO | Notes |
|----|------|---------|-----------|-------|
| Insert Image | pptx_insert_image (244-266) | InsertImage (359-394) | string | Append to shapeTree; GetMaxShapeId(); validate image file |
| Update Text | pptx_update_slide_data (94-150) | UpdateSlideData (190-200) | SlideDataUpdateResult | Locate by name/index; preserve formatting (complex) |
| Batch Update | pptx_batch_update (158-208) | BatchUpdate (202-235) | BatchUpdateResult | Single open/save; collect modified parts |
| Add Slide | pptx_add_slide (54-71) | AddSlide (113-171) | string | Create new SlidePart; link to layout |

**Insert Table** ≈ **Insert Image** (append to shapeTree, assign ID)
**Update Table** ≈ **Update Text** (locate by name) + **Insert Image** (GraphicFrame handling)

