# Skill: MCP Tool Descriptions

**Category:** MCP SDK / Server Metadata  
**Domain:** Tool discovery, runtime metadata, XML documentation  
**Maturity:** Established  
**Complexity:** Low

---

## Scope

Use this pattern when MCP tools are present in source with XML `<summary>` comments, but clients or startup logs report that the tools have no description.

---

## Core Pattern

### Generate the XML documentation file

In this codebase, ModelContextProtocol populates advertised tool descriptions from the compiled XML documentation output, not just from source comments alone.

Add this to the tool project file:

```xml
<GenerateDocumentationFile>true</GenerateDocumentationFile>
```

Current location:

```xml
src/PptxTools/PptxTools.csproj
```

---

## Verification Pattern

1. Build the tool project or solution.
2. Confirm the output folder contains `PptxTools.xml`.
3. Reflect or list MCP tools and verify `ProtocolTool.Description` is populated.
4. If a warning was reported from `dotnet run`, force a fresh **Debug** rebuild and re-check `src/PptxTools/bin/Debug/net10.0/PptxTools.xml` before touching source comments again.

Example affected files:

- `src/PptxTools/Tools/PptxTools.ManageMedia.cs`
- `src/PptxTools/Tools/PptxTools.ManageSlides.cs`
- `src/PptxTools/Tools/PptxTools.Optimization.cs`
- `src/PptxTools/Tools/PptxTools.Hyperlinks.cs`

---

## Anti-Pattern

- **Relying on source XML comments alone** — the comments can exist in code while runtime MCP metadata still exposes an empty description if the XML doc sidecar is not emitted.
- **Trusting stale local build output** — a warning can survive on a machine that is launching an older Debug build without the regenerated `PptxTools.xml`, even when the current branch source and fresh build are correct.
