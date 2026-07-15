# Phase 11C — Text Rendering Engine (Milestone 4)

**Date:** July 9, 2026  
**Status:** ✅ Complete — Build: 0 errors, 15 warnings  
**Location:** `ExcelAPI/ExcelAPI/Rendering/`

---

## Objective

Implement the production Text Rendering Engine for FormLess. The text engine renders cell text using SkiaSharp, using computed cell rectangles from GeometryBuilder — never approximate offsets.

---

## Current Rendering Pipeline

```
Workbook
  ↓
OpenXmlParser
  ↓
GeometryBuilder
  ↓
RenderingContext
  ↓
FillEngine        (Layer 1 — backgrounds)
  ↓
GridlineLayer     (Layer 2 — gridlines)
  ↓
BorderEngine      (Layer 3 — borders)
  ↓
TextEngine        (Layer 4 — text)  ← NEW
  ↓
ImageEngine       (future — Layer 5)
  ↓
PDF / PNG
```

The TextEngine implements `IRenderLayer` and is automatically executed by `RendererCoordinator` through the `IEnumerable<IRenderLayer>` DI pattern. No coordinator changes were required.

---

## Files Created (3 new)

### `Rendering/FontResolver.cs`

Resolves SkiaSharp typefaces from font names with a multi-level fallback chain. Never crashes if a font is unavailable — always returns a usable typeface.

**Fallback chain by font category:**

| Requested Font | Fallback Chain |
|:---------------|:---------------|
| Aptos Narrow (new Office default) | Calibri → Segoe UI → Arial → system default |
| Calibri | Segoe UI → Arial → Calibri |
| Meiryo | Yu Gothic → Segoe UI → Arial |
| Yu Gothic / MS Gothic | Meiryo → Segoe UI → Arial |
| MS Mincho | Yu Mincho → Meiryo → Segoe UI |
| Times New Roman / Serif | Times New Roman → Georgia → Segoe UI |
| Courier New / Monospace | Courier New → Consolas → Segoe UI |
| Generic sans-serif | Segoe UI → Arial → Calibri |

### `Rendering/TextLayoutEngine.cs`

Computes text layout from cell data. Produces `TextDrawCommand` objects with pixel-accurate positioning. No rendering here — only layout calculation.

**TextDrawCommand properties:**

| Property | Description |
|:---------|:------------|
| `Text` | Text content to render |
| `X, Y, Width, Height` | Pixel layout area |
| `HorizontalAlignment` | "left", "center", "right", "justify", "fill", "centerContinuous", "distributed" |
| `VerticalAlignment` | "top", "center", "bottom", "justify", "distributed" |
| `WrapText` | Whether text wraps to fit width |
| `RotationDegrees` | Text rotation (0, 90, -90, 180) |
| `Typeface` | Resolved SkiaSharp typeface |
| `FontSizePt` | Font size in points |
| `FontColor` | Resolved SKColor |
| `Underline` / `Strikeout` | Text decoration flags |
| `ClipRect` | Clipping rectangle (cell/merge bounds) |
| `Indent` / `IndentPixels` | Indentation level |

**Alignment resolution:** Excel-style "general" alignment = text left, numbers right.

### `Rendering/TextEngine.cs`

Implements `IRenderLayer` — renders cell text via SkiaSharp with proper pipeline integration.

**Rendering order within this layer:**

1. **Skip non-anchor merge cells** — only the top-left cell of each merged range renders (matches Excel behavior)
2. **Save canvas state** — `canvas.Save()` for clip/transform isolation
3. **Clip** — apply cell/merge bounds as clip rectangle (prevents text spill)
4. **Rotate** — apply rotation transform around cell center (0, 90, -90, 180 degrees)
5. **Align** — compute X/Y position based on horizontal + vertical alignment
6. **Wrap** — binary-search word wrapping for `WrapText` cells (handles newlines)
7. **Draw** — render each line with SkiaSharp `DrawText`
8. **Decorate** — draw underline and strikethrough as proportional lines
9. **Restore** — `canvas.Restore()` to undo clip and transform

---

## Files Modified (3 changed)

### `Rendering/WorkbookModel.cs`

Added text rendering fields to `RenderCell`:

```csharp
public bool Underline { get; set; }      // NEW
public bool Strikeout { get; set; }      // NEW
```

### `Rendering/OpenXmlParser.cs`

Parser now reads text decoration from styles.xml in `ApplyCellStyle`:

```csharp
rc.Underline = font.Underline != null;     // Any underline type = true
rc.Strikeout = font.Strike?.Val ?? false;  // OpenXml "Strike" property
```

### `Program.cs`

Registered three new services in DI:

```csharp
// Register Text Rendering Engine (Phase 11C / M4)
builder.Services.AddSingleton<FontResolver>();
builder.Services.AddSingleton<TextLayoutEngine>();
builder.Services.AddSingleton<TextEngine>();
builder.Services.AddSingleton<IRenderLayer, TextEngine>();  // Layer 4, after borders
```

---

## Code Review — Issues Fixed

| Issue | Fix |
|:------|:----|
| ParseColor duplication in TextLayoutEngine | Delegated to `FillEngine.ParseColor()` |
| Unnecessary try/catch in FontResolver | Removed — `SKTypeface.FromFamilyName` never throws |
| IsCjkFont dead code | Removed (not needed for initial implementation) |

## Known Limitations

| Issue | Status | Notes |
|:------|:-------|:------|
| ShrinkToFit | ❌ Not implemented | Text may overflow if cell is smaller than content |
| "Fill" alignment | ⚠️ Approximate | Treated as right-aligned; Excel repeats characters |
| Justify/Distributed vertical | ⚠️ Partial | Falls through to "top" — acceptable for single-line text |
| Theme color resolution | ❌ Not implemented | Existing limitation from Phase 11A — returns null for theme colors |
| CJK font metrics | ⚠️ Basic | No specialized CJK metrics — generally acceptable fallback |

---

## Build Validation

| Metric | Result |
|:-------|:------:|
| Build errors | **0** ✅ |
| Build warnings | **15** (unrelated to new code) |
| Breaking API changes | **None** — DI resolves new services automatically |
| IRenderLayer integration | ✅ Auto-injected via `IEnumerable<IRenderLayer>` |
| Frontend compatibility | ✅ No changes to Next.js or API contracts |

---

## File Inventory

```
ExcelAPI/ExcelAPI/Rendering/
├── FontResolver.cs          (NEW)  ~100 lines
├── TextLayoutEngine.cs      (NEW)  ~120 lines
├── TextEngine.cs            (NEW)  ~280 lines
├── WorkbookModel.cs         (MOD)  +2 properties
├── OpenXmlParser.cs         (MOD)  +2 lines in ApplyCellStyle
└── ... (existing files unchanged)

ExcelAPI/Program.cs          (MOD)  +5 lines DI registration
```
