# EMU Calculator & Reference Guide

English Metric Units (EMUs) are the standard unit of measurement in OpenXML/PowerPoint. Understanding EMU conversion is essential when positioning and sizing shapes, images, and tables in presentations.

---

## Quick Reference

**914,400 EMUs = 1 inch = 2.54 centimeters**

### Common Conversions

| Inches | EMUs | Centimeters | Use Case |
|--------|------|-------------|----------|
| 0.5 in | 457,200 | 1.27 cm | Small margin, half-inch offset |
| 1.0 in | 914,400 | 2.54 cm | Standard spacing, 1-inch inset |
| 1.5 in | 1,371,600 | 3.81 cm | Table positioning, top margin |
| 2.0 in | 1,828,800 | 5.08 cm | Medium element spacing |
| 3.0 in | 2,743,200 | 7.62 cm | Large spacing, centered positioning |
| 5.0 in | 4,572,000 | 12.7 cm | Quarter-slide width |
| 10.0 in | 9,144,000 | 25.4 cm | Full-slide width (standard) |

---

## Standard Slide Dimensions

### 16:9 Widescreen (Default)

| Dimension | EMUs | Inches | Centimeters |
|-----------|------|--------|-------------|
| **Width** | 9,144,000 | 10.0 in | 25.4 cm |
| **Height** | 5,143,500 | 5.625 in | 14.29 cm |
| **Area** | 47,040,000,000 EMU² | 56.25 in² | 365.9 cm² |

### 4:3 Standard (Legacy)

| Dimension | EMUs | Inches | Centimeters |
|-----------|------|--------|-------------|
| **Width** | 9,144,000 | 10.0 in | 25.4 cm |
| **Height** | 6,858,000 | 7.5 in | 19.05 cm |
| **Area** | 62,739,120,000 EMU² | 75.0 in² | 483.9 cm² |

---

## Manual Conversion Formula

To convert any measurement:

```
EMUs = inches × 914,400
inches = EMUs ÷ 914,400
centimeters = inches × 2.54
```

### Examples

**Convert 2.5 inches to EMUs:**
```
2.5 × 914,400 = 2,286,000 EMUs
```

**Convert 3,000,000 EMUs to inches:**
```
3,000,000 ÷ 914,400 ≈ 3.28 inches
```

**Convert 5 centimeters to EMUs:**
```
5 ÷ 2.54 ≈ 1.97 inches
1.97 × 914,400 ≈ 1,801,368 EMUs
```

---

## Common Shape Positioning Scenarios

### Title at Top (1.5" from top, full width minus 0.5" margins)

```
Position X: 457,200 EMUs (0.5")
Position Y: 1,371,600 EMUs (1.5")
Width: 8,229,600 EMUs (9.0")
Height: 914,400 EMUs (1.0")
```

### Content Box Centered (3 inches from top, 1 inch inset)

```
Position X: 914,400 EMUs (1.0")
Position Y: 2,743,200 EMUs (3.0")
Width: 7,315,200 EMUs (8.0")
Height: 2,286,000 EMUs (2.5")
```

### Small Icon (0.5" square, bottom-right corner)

```
Position X: 8,229,600 EMUs (9.0")
Position Y: 4,571,100 EMUs (5.0")
Width: 457,200 EMUs (0.5")
Height: 457,200 EMUs (0.5")
```

### Full-Bleed Image (no margins)

```
Position X: 0 EMUs
Position Y: 0 EMUs
Width: 9,144,000 EMUs (10.0")
Height: 5,143,500 EMUs (5.625" for 16:9)
```

---

## Precision & Rounding

PowerPoint and Office tools internally use EMUs with 32-bit integer precision. For practical purposes:

- **Minimum visible offset:** ~20,000 EMUs (0.022")
- **Recommended precision:** Round to nearest 1,000 EMUs (0.001")
- **Practical alignment grid:** 457,200 EMUs (0.5") or 228,600 EMUs (0.25")

When specifying positions and sizes in pptx-tools tools, use integer EMU values. Floating-point values are accepted but will be rounded internally.

---

## Tools That Use EMUs

| Tool | EMU Parameters | Notes |
|------|---|---|
| `pptx_insert_table` | `x`, `y`, `width`, `height` | Position and size the table on the slide. Defaults: `x=1,371,600` (1.5"), `y=1,371,600` (1.5") |
| `pptx_insert_image` | `x`, `y`, `width`, `height` | Optional; defaults to fill shape or use layout geometry |
| `pptx_manage_slides` (Duplicate) | Shape positioning inherited from layout | No EMU input; uses existing shape geometry |

---

## Tips for Consistent Layouts

1. **Use layout placeholders first** — most shapes inherit position/size from slide layout; avoid manual EMU positioning
2. **Reference slide dimensions** — always validate 16:9 vs. 4:3 before setting absolute positions
3. **Leave margins** — human-readable layouts typically have 0.5"–1.0" margins on all sides
4. **Test in PowerPoint** — verify EMU-based shapes render as expected by opening the generated .pptx in Microsoft Office
5. **Align to grid** — round positions to nearest 457,200 EMUs (0.5" grid) for consistent visual appearance

---

## Troubleshooting EMU Issues

### Shape Appears Off-Screen

**Cause:** Position X or Y exceeds slide dimensions.
**Solution:** Verify X < 9,144,000 and Y < 5,143,500 (for 16:9). Reduce position values if needed.

### Shape Too Large or Small

**Cause:** Width or height converted incorrectly.
**Solution:** Double-check conversion formula. Use the Quick Reference table above.

### EMUs Show as Decimal

**Cause:** Tool accepted floating-point EMUs.
**Solution:** Round to nearest integer before passing to tool. pptx-tools will round internally, but explicit rounding is clearer.

### Inconsistent Sizing After Update

**Cause:** Different shape inherits different scale factor from layout.
**Solution:** Use layout placeholders consistently; avoid mixing manual EMU positioning with layout-inherited shapes.

---

## Related Resources

- [TOOL_REFERENCE.md](TOOL_REFERENCE.md) — Complete tool parameter documentation
- [QUICKSTART.md](QUICKSTART.md) — Setup and first-time usage
- Microsoft OpenXML documentation: [Units in DrawingML](https://docs.microsoft.com/en-us/office/open-xml/articles/working_with_drawingml) (external)
