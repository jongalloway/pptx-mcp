# Markdown export output defaults

- Tool: `pptx_export_markdown`
- Decision: Write markdown to the requested output path, or default to the source deck path with a `.md` extension.
- Images: Extract embedded images to a sibling `<markdown-base>_images` folder and reference them with relative markdown paths.
- Scope: Phase 1 export excludes speaker notes and focuses on visible slide content.

This keeps exported markdown portable on disk while still returning the markdown body directly to MCP callers.
