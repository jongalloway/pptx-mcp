# Shape selection strategy for `pptx_update_slide_data`

## Status
Implemented for issue #19.

## Decision
1. **Primary selector: `shapeName`**
   - Match against the slide shape's OpenXML name (`p:cNvPr/@name`) with case-insensitive exact matching.
   - This is the best UX for agents because `pptx_get_slide_content` already exposes stable names like `Title 1`, `Content Placeholder 2`, or template-specific metric box names.

2. **Fallback selector: `placeholderIndex`**
   - Interpret it as a **zero-based index across text-capable slide shapes (`p:sp`) in slide order**.
   - This works when names are missing, generic, or not trustworthy in a template.
   - If both selectors are provided, name lookup wins; if the name is not found and an index is present, fall back to the index.

3. **Failure behavior**
   - Duplicate shape-name matches fail fast instead of picking one arbitrarily.
   - Missing-name and out-of-range-index failures should return the available shape indexes and names so an agent can recover in the next call.

## Deferred options
- **Placeholder type (`title`, `body`, `subtitle`)**: not implemented in v1 because multiple shapes on a slide can share the same semantic type, which makes updates ambiguous without additional rules.
- **Template markers like `{{metric}}`**: useful later, but they solve a different templating workflow than direct shape targeting.

## Working guidance
Agents should call `pptx_get_slide_content` first, inspect `Name` and `ShapeType`, then call `pptx_update_slide_data` with `shapeName` whenever possible. Use `placeholderIndex` only as a deterministic fallback when the deck does not expose meaningful names.
