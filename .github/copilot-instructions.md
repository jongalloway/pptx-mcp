# GitHub Copilot Instructions for pptx-tools

## Project Overview

This project is an **MCP (Model Context Protocol) server** that reads and modifies PowerPoint (.pptx) files using the OpenXML SDK. It exposes PPTX operations as MCP tools that AI assistants can invoke to inspect, analyze, and edit presentation files.

**Important**: This package is designed exclusively as an **MCP server** for AI assistants. It is not intended for use as a library or for programmatic consumption in other .NET applications.

## Key Technologies

- **.NET 10.0**: Latest version of .NET
- **Model Context Protocol SDK**: `ModelContextProtocol` v1.1.0
- **DocumentFormat.OpenXml**: v3.3.0 — primary implementation surface for PPTX operations
- **Microsoft.Extensions.Hosting**: Application lifecycle management
- **Stdio Transport**: Communication via standard input/output

## Architecture

- `src/PptxTools/` — Main MCP server project
  - `Program.cs` — Host builder, MCP server registration
  - `Tools/` — MCP tool methods (marked with `[McpServerToolType]` / `[McpServerTool]`)
  - `Services/` — Business logic (e.g., `PresentationService` for OpenXML operations)
  - `Models/` — Data transfer objects for tool responses
- `tests/PptxTools.Tests/` — xUnit v3 tests using Microsoft Testing Platform

## Code Style and Conventions

### Naming
- PascalCase for classes, methods, public members
- camelCase for local variables and parameters
- Tool method naming: `{Noun}{Verb}` pattern (e.g., `SlideList`, `SlideGetText`)

### MCP Tool Patterns
- Mark tool classes with `[McpServerToolType]`
- Mark tool methods with `[McpServerTool]` and make them `partial`
- Use XML doc comments (`/// <summary>`, `/// <param>`) for tool and parameter descriptions — the MCP SDK generates `Description` metadata from these
- Use nullable types for optional parameters with default values
- Return meaningful string results or structured JSON

### OpenXML Patterns
- **PowerPoint compatibility is the real success criterion.** A file can pass `OpenXmlValidator` and still fail to open in PowerPoint.
- Preserve package structure and relationship invariants when modifying PPTX files
- Use `DocumentFormat.OpenXml` as the primary implementation surface
- When troubleshooting, inspect package structure, relationships, and content types before changing slide content logic
- Prefer `Path.Join(...)` over `Path.Combine(...)` for path handling

### Reference Projects
- **jongalloway/dotnet-mcp**: Reference for MCP SDK patterns, tool registration, conformance tests, CI/CD
- **jongalloway/MarpToPptx**: Reference for OpenXML SDK patterns, PPTX generation, compatibility testing

## Building and Testing

```bash
# Build
dotnet build PptxTools.slnx --configuration Release

# Test (uses Microsoft Testing Platform via global.json)
dotnet test --solution PptxTools.slnx --configuration Release --no-build

# Targeted test run (xUnit v3 under MTP)
dotnet test --project tests/PptxTools.Tests/PptxTools.Tests.csproj -c Release -- --filter-method "*SpecificTestMethod"
```

Do not assume VSTest-style `--filter` works here; under MTP/xUnit v3, use `--filter-method`, `--filter-class`, or other runner-supported options after `--`.

## Adding New MCP Tools

1. Add a method to the appropriate class in `Tools/`
2. Mark with `[McpServerToolType]` (class) and `[McpServerTool]` (method, `partial`)
3. Add XML doc comments for tool and all parameters
4. Implement business logic in `Services/` — keep tools thin
5. Add tests in `tests/PptxTools.Tests/`
6. Update README.md to list the new tool

## Documentation Guidelines

- Update README.md when adding new tools or features
- Do NOT create summary documents for individual changes
- Keep this instructions file focused and small
- Put specialized workflows into prompts or skills

## CI/CD

- GitHub Actions runs build + test on push to main and PRs
- Coverage collected via Microsoft.Testing.Extensions.CodeCoverage
- copilot-setup-steps.yml enables Copilot coding agent for the repo
