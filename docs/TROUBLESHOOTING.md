# Troubleshooting Guide

Common issues, their causes, and solutions when using pptx-tools.

---

## Installation & Setup Issues

### "dotnet: command not found"

**Cause:** .NET SDK is not installed or not in your PATH.

**Solution:**
1. Download [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
2. Install it following the official guide for your OS
3. Verify: Open a new terminal and run `dotnet --version`
4. Should print `10.x.x` or later

---

### "Failed to build PptxTools — missing dependencies"

**Cause:** NuGet packages failed to restore.

**Solution:**
1. Clear NuGet cache:
   ```bash
   dotnet nuget locals all --clear
   ```
2. Restore explicitly:
   ```bash
   dotnet restore PptxTools.slnx
   ```
3. Try building again:
   ```bash
   dotnet build PptxTools.slnx --configuration Release
   ```

---

### "MCP Server doesn't respond to tool calls"

**Cause:** The server isn't listening, or the MCP client isn't sending messages correctly.

**Check:**
1. Is the server running? Test with:
   ```bash
   dotnet run --project src/PptxTools/PptxTools.csproj --configuration Release
   ```
   Press Ctrl+C; if no error, the server ran. Start it again.

2. Is Claude Desktop (or your MCP client) configured correctly?
   - On macOS: `~/Library/Application\ Support/Claude/claude_desktop_config.json`
   - On Windows: `%APPDATA%\Claude\claude_desktop_config.json`
   - On Linux: `~/.config/Claude/claude_desktop_config.json`
   - Verify `"command"` and `"args"` point to the correct .NET binary path

3. Restart your MCP client after any configuration changes

**Solution:** See [QUICKSTART.md](QUICKSTART.md) for detailed setup steps.

---

## File & Path Issues

### "File not found: /path/to/presentation.pptx"

**Cause:** The file doesn't exist at the specified path, or the path is incorrect.

**Solution:**
1. Verify the path:
   - On Windows: Use backslashes or escape them: `C:\path\to\file.pptx` or `C:/path/to/file.pptx`
   - On macOS/Linux: Use forward slashes: `/path/to/file.pptx`
2. Check that the file actually exists: `ls /path/to/file.pptx` (macOS/Linux) or `dir C:\path\to\file.pptx` (Windows)
3. Verify you have read/write permissions on the file and its directory

---

### "Access denied — insufficient permissions"

**Cause:** The file is open in PowerPoint, or your user lacks permissions.

**Solution:**
1. **If the file is open in PowerPoint:** Close PowerPoint and the file completely. Try again.
2. **If the file is locked by another process:** Use `lsof` (macOS/Linux) or Task Manager (Windows) to find what's holding the file open
3. **If you lack permissions:** Run your terminal as Administrator (Windows) or use `sudo` (macOS/Linux)

---

### "Can't write to the output file"

**Cause:** The destination directory doesn't exist or lacks write permissions.

**Solution:**
1. Create the output directory: `mkdir -p /path/to/output`
2. Check permissions: `ls -l /path/to` (macOS/Linux) should show `w` for your user
3. On Windows, right-click the folder → Properties → Security → Edit → ensure your user has write permissions

---

## PowerPoint Format Issues

### "File is not a valid PowerPoint presentation"

**Cause:** The file isn't a real `.pptx`, or it's corrupted.

**Solution:**
1. Verify the file is actually PowerPoint format: `file /path/to/file.pptx` (macOS/Linux)
2. Try opening the file in Microsoft Office or PowerPoint. If it opens, the file is valid.
3. If opening fails in PowerPoint too, the file is likely corrupted. Use a backup or recreate the file.
4. Ensure the file extension is `.pptx` (not `.ppt`, `.odp`, or `.pdf`)

---

### "Presentation is corrupted and cannot be recovered"

**Cause:** The .pptx file's internal XML structure is invalid.

**Solution:**
1. Try extracting and re-packaging the .pptx:
   ```bash
   cd /path/to/file.pptx  # .pptx is a ZIP file
   unzip file.pptx -d extracted/
   # Inspect the XML in extracted/ppt/slides/
   # If you spot corruption, edit the XML manually
   cd extracted
   zip -r ../file-fixed.pptx .
   ```
2. Use an online PPTX validator to check the file structure
3. If the file is important, use PowerPoint's built-in "Repair" option:
   - In PowerPoint, go to File → Open → Right-click the file → Open and Repair

---

## Shape & Slide Issues

### "Shape 'XYZ' not found on slide 1"

**Cause:** The shape name is misspelled, or the shape doesn't exist on that slide.

**Solution:**
1. Discover available shapes:
   ```bash
   # Use the shape-map resource (see SHAPE_RESOLUTION_GUIDE.md)
   # Or call pptx_get_slide_content to inspect the slide
   ```
2. Check the spelling (shape names are case-sensitive: `"Title"` ≠ `"title"`)
3. Verify the slide number (1-based in most contexts, 0-based in others — check the tool docs)

---

### "Shape index 5 is out of range on slide 1"

**Cause:** The slide has fewer than 6 shapes.

**Solution:**
1. Shapes are 0-indexed. If the slide has 4 shapes, valid indices are 0, 1, 2, 3 (not 4, 5, 6).
2. Call `pptx_get_slide_content` to see how many shapes exist and which indices are valid.

---

### "Placeholder 'Body:2' not found"

**Cause:** The slide layout doesn't have a second body placeholder.

**Solution:**
1. Check how many body placeholders the layout actually has:
   ```bash
   # Call pptx_list_layouts and inspect the layout in PowerPoint
   # Or check the slide master view in PowerPoint
   ```
2. Use `"Body"` or `"Body:1"` if the layout only has one body placeholder
3. If you need multiple text areas, use `pptx_insert_table` or add additional text shapes manually in PowerPoint

---

### "New slide not appearing"

**Cause:** The slide was created but not visible in PowerPoint.

**Solution:**
1. Save the presentation after the tool call. pptx-tools saves automatically, but confirm the file's modification timestamp changed.
2. Close and reopen PowerPoint to refresh the view
3. Try scrolling in the slide sorter view to find the new slide
4. Verify the slide number was returned correctly by the tool

---

## Text & Content Issues

### "Text appears garbled or cut off"

**Cause:** Text encoding mismatch or placeholder too small for content.

**Solution:**
1. **Garbled text:** Verify the text is valid UTF-8. pptx-tools supports Unicode. Try simpler text without special characters.
2. **Cut off:** The placeholder shape might be too small. Try:
   - Shorter text
   - Smaller font (adjust in PowerPoint)
   - A larger placeholder (resize in PowerPoint or use a different layout)
3. **Word wrap:** PowerPoint wraps text automatically in placeholders. If text doesn't wrap, the shape might have `NoAutofit` enabled in PowerPoint.

---

### "Table content appears misaligned"

**Cause:** Column count mismatch or cell coordinates wrong.

**Solution:**
1. Verify the number of columns matches your data: `pptx_insert_table` requires all rows to have the same column count
2. Check row/column indices are 0-based in `pptx_update_table`
3. If a cell is empty, try providing a single space `" "` instead of empty string

---

### "Image appears as a broken link or placeholder"

**Cause:** Image file doesn't exist, or embedded image is corrupted.

**Solution:**
1. **For `pptx_insert_image` with a file path:** Verify the image file exists and is readable: `ls /path/to/image.png`
2. **For embedding:** Try a different image format (PNG, JPEG, SVG)
3. **For large images:** Compress the image first:
   ```bash
   # On macOS/Linux, use ImageMagick
   convert large.jpg -quality 85 -resize 50% compressed.jpg
   ```

---

## Chart & Data Issues

### "Chart data not updated"

**Cause:** The chart name/index is wrong, or the data format is invalid.

**Solution:**
1. Verify the chart name with `pptx_get_slide_content`
2. Ensure all series have the same number of values: `["A", "B", "C"]` values for 3 categories
3. For numeric charts, values must be numbers, not strings: `[100, 200, 150]` not `["100", "200", "150"]`
4. Check that the chart type matches your data (e.g., don't put text values in a scatter plot)

---

### "Chart type not supported"

**Cause:** You're trying to update a chart type that pptx-tools doesn't support.

**Solution:**
Supported types: Column, Bar, Line, Pie, Area, Scatter, and their 3D/Doughnut variants.

**Unsupported:** Bubble charts, stock charts, surface charts.

If you need an unsupported chart type, create it manually in PowerPoint and use `pptx_chart_data` to read/update its data.

---

## Layout & Template Issues

### "Layout 'Title and Content' not found"

**Cause:** The layout name is misspelled or doesn't exist in this presentation's template.

**Solution:**
1. List available layouts:
   ```bash
   # Call pptx_list_layouts to see all layouts in the presentation
   ```
2. Check the exact spelling (case-sensitive)
3. If the layout you need doesn't exist, create it in PowerPoint's slide master view, then try again

---

### "New slide inherits wrong layout"

**Cause:** You didn't specify a layout, or specified the wrong one.

**Solution:**
1. Always call `pptx_list_layouts` to see available layouts
2. Use `pptx_add_slide_from_layout` instead of `pptx_add_slide` to control which layout is used
3. If the layout changes after creation, the layout name might be wrong or the file has been modified

---

## Performance Issues

### "Tool call times out (30+ seconds)"

**Cause:** The presentation is very large or the operation is disk I/O bound.

**Solution:**
1. Close other applications to free up disk I/O
2. Try with a smaller presentation first to verify the tool works
3. For very large presentations (>100 MB), consider splitting into smaller files
4. If the timeout is from your MCP client, increase the client's timeout setting

---

### "Memory usage is very high"

**Cause:** Large presentations or many images in memory.

**Solution:**
1. Use batch operations (`pptx_batch_update`) instead of multiple single updates
2. For image-heavy presentations, ensure images are compressed (`pptx_optimize_images`)
3. Close and restart the MCP server between large operations if memory doesn't release

---

## Testing Issues

### "Tests fail: 'OpenXML validation error'"

**Cause:** The generated PPTX violates OpenXML schema requirements.

**Solution:**
1. Check the error message for the specific violation (e.g., "missing required element `p:cSld`")
2. Review the tool's implementation or contact the maintainers
3. Validate your input data (e.g., table rows must have consistent column counts)

---

### "Test fixture files not found"

**Cause:** Test .pptx files in `tests/fixtures/` are missing.

**Solution:**
1. Verify the test fixture files exist: `ls tests/fixtures/`
2. If missing, run the test suite once to auto-generate fixtures:
   ```bash
   dotnet test --solution PptxTools.slnx --configuration Release
   ```
3. If auto-generation doesn't work, the fixtures are created by `TestPptxHelper` — check the helper's output directory

---

## Building & CI Issues

### "Build fails: Restore incomplete"

**Cause:** NuGet packages failed to download.

**Solution:**
1. Check your internet connection
2. Clear the NuGet cache:
   ```bash
   dotnet nuget locals all --clear
   ```
3. Try restoring explicitly:
   ```bash
   dotnet restore PptxTools.slnx
   ```

---

### "Tests fail in CI but pass locally"

**Cause:** Environment differences (OS, .NET version, file system).

**Solution:**
1. Verify CI .NET version matches your local `dotnet --version`
2. Check for hard-coded paths in tests (should be relative)
3. Ensure test fixtures don't rely on OS-specific paths (use `/` or `Path.Combine`)
4. Try running tests with `--configuration Release` locally to match CI

---

## Getting Help

### Check the Documentation

1. [QUICKSTART.md](QUICKSTART.md) — Setup and first-time usage
2. [TOOL_REFERENCE.md](TOOL_REFERENCE.md) — Full parameter documentation for all tools
3. [EMU_GUIDE.md](EMU_GUIDE.md) — Understanding and converting EMU coordinates
4. [SHAPE_RESOLUTION_GUIDE.md](SHAPE_RESOLUTION_GUIDE.md) — Targeting shapes by name, index, or placeholder
5. [EXAMPLES.md](EXAMPLES.md) — Real-world usage examples

### Report an Issue

If you encounter a bug not listed here:

1. Create a [GitHub issue](https://github.com/jongalloway/pptx-tools/issues/new)
2. Include:
   - Your .NET version (`dotnet --version`)
   - A minimal reproduction case (smallest .pptx + tool call that reproduces the issue)
   - The exact error message and stack trace
   - Your OS (Windows/macOS/Linux)

### Contributing a Fix

See [CONTRIBUTING.md](CONTRIBUTING.md) for how to submit a pull request.

---

## Related Resources

- [TOOL_REFERENCE.md](TOOL_REFERENCE.md) — Complete tool documentation
- [SHAPE_RESOLUTION_GUIDE.md](SHAPE_RESOLUTION_GUIDE.md) — Targeting shapes
- [EMU_GUIDE.md](EMU_GUIDE.md) — Coordinates and sizing
- [CONTRIBUTING.md](CONTRIBUTING.md) — Contributing & development setup
