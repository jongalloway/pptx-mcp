# Contributing to pptx-tools

Thank you for your interest in contributing to pptx-tools! This guide explains the development workflow, how to add new tools, and the test and CI/CD process.

---

## Prerequisites

- **[.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)** — verify with `dotnet --version` (must be 10.x.x or later)
- **[Git](https://git-scm.com/downloads)**
- A GitHub account (for opening issues and PRs)

---

## Development Setup

### 1. Clone the Repository

```bash
git clone https://github.com/jongalloway/pptx-tools.git
cd pptx-tools
```

### 2. Build the Project

```bash
dotnet build PptxTools.slnx --configuration Release
```

Verify the build succeeds with no errors.

### 3. Run the Test Suite

```bash
dotnet test --solution PptxTools.slnx --configuration Release
```

All tests should pass. Expected output: `1227 passed`.

### 4. Verify the Server Runs

```bash
dotnet run --project src/PptxTools/PptxTools.csproj --configuration Release
```

The server starts and waits for MCP messages on stdin. Press Ctrl+C to stop.

---

## Project Structure

```
pptx-tools/
├── src/PptxTools/                    # Main MCP server
│   ├── Commands/                     # Command-line argument handling
│   ├── Completions/                  # MCP argument auto-completion
│   ├── Models/                       # Data classes (OperationResult, etc.)
│   ├── Services/                     # Business logic (PresentationService, etc.)
│   ├── Tools/                        # MCP tool implementations
│   ├── Resources/                    # MCP resources
│   ├── Prompts/                      # MCP prompts
│   ├── Program.cs                    # Entry point
│   └── PptxTools.csproj              # Project file
│
├── tests/PptxTools.Tests/            # xUnit v3 tests (Microsoft Testing Platform)
│   ├── [Tool]Tests.cs                # Tool-level tests
│   ├── [Service]Tests.cs             # Service-level tests
│   ├── TestPptxHelper.cs             # Shared test fixture builder
│   ├── PptxTestBase.cs               # Base class for file-based tests
│   └── PptxTools.Tests.csproj        # Test project file
│
├── docs/                             # User & contributor documentation
│   ├── EMU_GUIDE.md                  # EMU conversion reference
│   ├── SHAPE_RESOLUTION_GUIDE.md     # Targeting shapes
│   ├── TROUBLESHOOTING.md            # Common issues & solutions
│   ├── CONTRIBUTING.md               # This file
│   ├── TOOL_REFERENCE.md             # Complete tool parameter docs
│   ├── QUICKSTART.md                 # Setup for users
│   ├── EXAMPLES.md                   # Usage examples
│   ├── MULTI_SOURCE_COMPOSITION.md   # Multi-MCP workflows
│   └── PRD.md                        # Architecture & design
│
├── .github/workflows/                # CI/CD automation
│   ├── build.yml                     # Build & test on every PR
│   └── ...                           # Squad orchestration workflows
│
├── PptxTools.slnx                    # Solution file
└── README.md                         # Project overview
```

---

## Architecture Overview

### Models → Services → Tools → MCP Server

**Models** (`src/PptxTools/Models/`)
- Data classes like `OperationResult`, `SlideInfo`, `ShapeInfo`
- No business logic, just data structures

**Services** (`src/PptxTools/Services/`)
- `PresentationService` (partial class with separate files per feature):
  - `PresentationService.Slides.cs` — slide operations
  - `PresentationService.Text.cs` — text manipulation
  - `PresentationService.Images.cs` — image insertion & management
  - ... etc.
- Pure business logic with no MCP dependencies
- Responsible for all OpenXML manipulation
- Each public method returns a strongly-typed result

**Tools** (`src/PptxTools/Tools/`)
- MCP tool implementations: `PptxTools.Slides.cs`, `PptxTools.Text.cs`, etc.
- Thin wrappers that call service methods
- Handle input validation and MCP-specific concerns
- Each tool method has `[Tool(...)]` attribute for MCP metadata

**MCP Server** (`src/PptxTools/Program.cs`)
- Wires up tools, resources, and prompts
- Handles stdio transport and message serialization

### Why This Structure?

- **Testability:** Service layer can be tested without MCP infrastructure
- **Reusability:** Services can be used by CLI tools, other projects, etc.
- **Clarity:** Business logic is separate from MCP plumbing
- **Maintainability:** Each concern has its own file and responsibility

---

## Adding a New Tool

### Example: Add `pptx_rotate_shape`

#### Step 1: Implement the Service Method

File: `src/PptxTools/Services/PresentationService.Shapes.cs` (or create if it doesn't exist)

```csharp
public OperationResult RotateShape(string filePath, int slideNumber, string shapeName, int degrees)
{
    try
    {
        var result = OpenPresentation(filePath);
        if (!result.Success)
            return result;

        var slide = GetSlide(result.Presentation, slideNumber);
        if (slide == null)
            return OperationResult.Failure($"Slide {slideNumber} not found.");

        var shape = slide.Descendants<P.Shape>()
            .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.Name == shapeName);

        if (shape == null)
            return OperationResult.Failure($"Shape '{shapeName}' not found on slide {slideNumber}.");

        // Apply rotation (rotation is in 60,000ths of a degree)
        var rot = shape.ShapeProperties?.Transform2D?.Rotation;
        if (rot == null)
            return OperationResult.Failure("Shape has no transform properties.");

        rot.Val = (int)((degrees % 360) * 60000);

        result.Presentation?.Save();
        return OperationResult.Success($"Rotated '{shapeName}' by {degrees}°.");
    }
    catch (Exception ex)
    {
        return OperationResult.Failure($"Error: {ex.Message}");
    }
}
```

#### Step 2: Add the MCP Tool

File: `src/PptxTools/Tools/PptxTools.Shapes.cs` (or create if it doesn't exist)

```csharp
public partial class PptxTools
{
    [Tool("Rotate a shape on a slide", "Rotates a shape by the specified degrees (0-360).")]
    public RotateShapeResult RotateShape(
        [ToolParameter("Absolute or relative path to the .pptx file")] string filePath,
        [ToolParameter("1-based slide number")] int slideNumber,
        [ToolParameter("Name of the shape to rotate")] string shapeName,
        [ToolParameter("Degrees to rotate (0-360)")] int degrees)
    {
        var serviceResult = _presentationService.RotateShape(filePath, slideNumber, shapeName, degrees);
        return new RotateShapeResult
        {
            Success = serviceResult.Success,
            Message = serviceResult.Message
        };
    }
}
```

File: `src/PptxTools/Models/RotateShapeResult.cs`

```csharp
public class RotateShapeResult
{
    public bool Success { get; set; }
    public string Message { get; set; }
}
```

#### Step 3: Add Tests

File: `tests/PptxTools.Tests/ShapeTests.cs` (or `RotateShapeTests.cs`)

```csharp
public class RotateShapeTests : PptxTestBase
{
    [Fact]
    public void RotateShape_ByName_Succeeds()
    {
        // Arrange
        var presentation = CreateTestPresentation();
        var result = _service.RotateShape(presentation.FilePath, 1, "TestShape", 45);

        // Act & Assert
        Assert.True(result.Success);
        Assert.Contains("45°", result.Message);
    }

    [Fact]
    public void RotateShape_InvalidShape_Returns_Failure()
    {
        // Arrange
        var presentation = CreateTestPresentation();

        // Act
        var result = _service.RotateShape(presentation.FilePath, 1, "NonExistent", 45);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.Message);
    }
}
```

#### Step 4: Update Documentation

Add to `docs/TOOL_REFERENCE.md`:

```markdown
## pptx_rotate_shape

**Description:** Rotate a shape on a slide.

### Parameters

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | ✅ Required | Absolute or relative path to the .pptx file. |
| `slideNumber` | integer | ✅ Required | 1-based slide number. |
| `shapeName` | string | ✅ Required | Name of the shape to rotate. |
| `degrees` | integer | ✅ Required | Rotation amount in degrees (0–360). |

### Returns

```json
{
  "success": true,
  "message": "Rotated 'shape1' by 45°."
}
```

### Example

```json
{
  "name": "pptx_rotate_shape",
  "arguments": {
    "filePath": "/presentations/diagram.pptx",
    "slideNumber": 1,
    "shapeName": "Arrow",
    "degrees": 90
  }
}
```
```

#### Step 5: Run Tests and Build

```bash
dotnet build PptxTools.slnx --configuration Release
dotnet test --solution PptxTools.slnx --configuration Release
```

All tests should pass, including your new ones.

---

## Test Patterns

### Service Tests (Unit Tests)

Service tests exercise business logic in isolation. They use `TestPptxHelper` to create realistic fixtures:

```csharp
public class UpdateTextTests
{
    private readonly PresentationService _service = new();

    [Fact]
    public void UpdateText_ByName_PreservesFormatting()
    {
        // Arrange
        var filePath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.pptx");
        TestPptxHelper.CreatePresentation(filePath, new[]
        {
            new TestSlideDefinition
            {
                TitleText = "Original Title",
                TextShapes = new[]
                {
                    new TestTextShapeDefinition
                    {
                        PlaceholderType = PlaceholderValues.Body,
                        Paragraphs = new[] { "Bullet 1", "Bullet 2" }
                    }
                }
            }
        });

        // Act
        var result = _service.UpdateText(filePath, 1, "Title", "New Title");

        // Assert
        Assert.True(result.Success);
        
        // Verify the change persisted
        var content = _service.GetSlideContent(filePath, 1);
        Assert.Contains("New Title", content.ShapesJson);

        File.Delete(filePath);
    }
}
```

### Tool Tests (Integration Tests)

Tool tests call the MCP tool methods and verify the MCP result format:

```csharp
public class UpdateTextToolTests
{
    private readonly PptxTools _tools = new();

    [Fact]
    public void UpdateText_Tool_ReturnsCorrectFormat()
    {
        // Arrange
        var filePath = CreateTestPresentation();

        // Act
        var result = _tools.UpdateText(filePath, 1, "Title", "New Title");

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Message);

        File.Delete(filePath);
    }
}
```

### Best Practices

1. **Name tests clearly:** `WhatYouAreTesting_InputCondition_ExpectedOutcome`
   - ✅ `UpdateText_ByName_PreservesFormatting`
   - ❌ `TestUpdate`

2. **Use `TestPptxHelper` for fixtures:** Don't hand-build OpenXML. Use the helper:
   ```csharp
   TestPptxHelper.CreatePresentation(filePath, slides);
   ```

3. **Clean up test files:** Delete temp .pptx files in test cleanup:
   ```csharp
   File.Delete(filePath);
   ```

4. **Test edge cases:** Invalid input, missing data, boundary conditions
   - Empty presentations
   - Missing shapes
   - Out-of-range indices
   - Unicode/special characters

5. **Verify side effects:** If a tool modifies a file, verify the change persists by re-reading

---

## Building & Testing Commands

### Build

```bash
# Build with Release configuration (optimized)
dotnet build PptxTools.slnx --configuration Release

# Build and report verbose output
dotnet build PptxTools.slnx --configuration Release --verbosity detailed

# Clean build (delete bin/obj first)
dotnet clean PptxTools.slnx
dotnet build PptxTools.slnx --configuration Release
```

### Test

```bash
# Run all tests
dotnet test --solution PptxTools.slnx --configuration Release

# Run a specific test class
dotnet test --solution PptxTools.slnx --configuration Release --filter-method UpdateTextTests

# Run a specific test method
dotnet test --solution PptxTools.slnx --configuration Release --filter-method UpdateTextTests.UpdateText_ByName_PreservesFormatting

# Run with code coverage
dotnet test --solution PptxTools.slnx --configuration Release -- --coverage --coverage-output-format cobertura
```

### Run the Server Locally

```bash
# Start the server (waits for MCP messages on stdin)
dotnet run --project src/PptxTools/PptxTools.csproj --configuration Release

# Or run the built binary directly
src/PptxTools/bin/Release/net10.0/PptxTools
```

---

## CI/CD Pipeline

pptx-tools uses GitHub Actions for continuous integration.

### Build Workflow

The `.github/workflows/build.yml` workflow:

1. **Setup**: Installs .NET 10
2. **Restore**: Downloads dependencies (`dotnet restore`)
3. **Build**: Compiles the project (`dotnet build --no-restore`)
4. **Test**: Runs all tests with code coverage (`dotnet test --no-build -- --coverage`)
5. **Coverage**: Uploads coverage report as artifact
6. **Summary**: Writes workflow summary to GitHub

### Workflow Triggers

The build runs on:
- **Every push to `main`** — verify main is always buildable
- **Every pull request to `main`** — verify changes don't break tests
- **Manual trigger** — `workflow_dispatch` in GitHub Actions UI

### Status Checks

All checks must pass before a PR can merge:
- ✅ Build succeeds
- ✅ All tests pass
- ✅ No style violations (if linter enabled)

### Viewing CI Results

1. Open a PR on GitHub
2. Scroll to "Checks" section
3. Click "Details" next to the build status
4. View step outcomes and logs

---

## Code Style & Conventions

### C# Style

- **Naming:** `PascalCase` for classes/methods, `camelCase` for variables
- **Formatting:** Default Visual Studio formatting (.editorconfig defines rules)
- **Null handling:** Use `?.` safe navigation operator, avoid `if (x != null)` checks when possible
- **Async:** Not currently used (all operations are synchronous I/O on files)

### File Organization

- **One class per file** (unless nested or very small)
- **Partial classes for related methods** — `PresentationService.Slides.cs`, `PresentationService.Text.cs`, etc.
- **No region blocks** — use file organization instead

### Comments

- Only comment complex logic
- Avoid obvious comments: `// Increment i` is not helpful
- Document public methods: `/// <summary>` XML doc comments

### Tests

- Arrange-Act-Assert (AAA) pattern
- One assertion per test (when possible) — test name should describe the assertion
- Use descriptive test data, not generic `"Test"`

---

## Opening a Pull Request

### Before You Start

1. **Open an issue first** (if one doesn't exist) describing the feature or bug
2. **Discuss significant changes** with maintainers before implementing
3. **Check existing PRs** — don't duplicate work

### Creating Your PR

1. **Create a branch:**
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b fix/your-bug-fix-name
   ```

2. **Make your changes** following the patterns above

3. **Commit with clear messages:**
   ```bash
   git commit -m "Add pptx_rotate_shape tool

   - Implement RotateShape service method
   - Add MCP tool wrapper
   - Add 5 test cases covering edge cases
   - Update TOOL_REFERENCE.md with parameter docs"
   ```

4. **Push and open a PR:**
   ```bash
   git push origin feature/your-feature-name
   ```

5. **In the PR description:**
   - Reference the issue: "Closes #123"
   - Explain what you changed and why
   - Mention any breaking changes
   - List any new test fixtures or dependencies

### PR Review Process

1. Automated checks run (build, tests, coverage)
2. Maintainers review your code
3. Respond to feedback and push updates
4. Once approved, your PR is merged

---

## Common Tasks

### Adding a New Tool

See "Adding a New Tool" section above.

### Running Tests During Development

```bash
# Quick test run (no coverage)
dotnet test --solution PptxTools.slnx --configuration Release

# Watch mode (not currently set up, but you can manually re-run)
# Just run the command above again after making changes

# Test a specific class
dotnet test --solution PptxTools.slnx --configuration Release --filter-method ShapeTests
```

### Debugging Tests

1. Set a breakpoint in Visual Studio Code or Visual Studio
2. Run tests with the debugger:
   ```bash
   dotnet test --solution PptxTools.slnx --configuration Debug
   ```
3. The debugger will stop at your breakpoint

### Updating Documentation

- Edit `.md` files in `docs/`
- No build or test needed for docs-only changes
- Verify links are correct: `[Link text](path/to/file.md)`
- Run a local markdown linter if available (optional)

---

## Troubleshooting Development Issues

### "Tests fail locally but pass in CI"

This usually means environment differences. Check:
1. .NET version: `dotnet --version` (must be 10.x.x)
2. Working directory: Run tests from repo root
3. Temp files: Delete `bin/` and `obj/` directories and rebuild

### "Build fails with NuGet error"

```bash
# Clear NuGet cache
dotnet nuget locals all --clear

# Restore explicitly
dotnet restore PptxTools.slnx

# Try building again
dotnet build PptxTools.slnx --configuration Release
```

### "Can't run the server locally"

```bash
# Make sure you're in the repo root
cd /path/to/pptx-tools

# Build first
dotnet build PptxTools.slnx --configuration Release

# Then run
dotnet run --project src/PptxTools/PptxTools.csproj --configuration Release
```

---

## Questions?

- **General questions:** Open a GitHub discussion or issue
- **Bug reports:** Include your .NET version, OS, and a minimal reproduction case
- **Feature requests:** Describe the use case and why it's valuable

---

## Licensing

By contributing to pptx-tools, you agree that your contributions will be licensed under the same license as the project (see LICENSE file).

---

## Related Resources

- [Architecture & Design](docs/PRD.md)
- [Tool Reference](docs/TOOL_REFERENCE.md)
- [Troubleshooting](docs/TROUBLESHOOTING.md)
- [EMU Guide](docs/EMU_GUIDE.md)
- [Shape Resolution Guide](docs/SHAPE_RESOLUTION_GUIDE.md)
