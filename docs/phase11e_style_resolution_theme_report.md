# Phase 11E — Excel Style Resolution & Theme Engine

**Date:** July 9, 2026  
**Status:** ✅ Complete — Build: 0 errors, 17 warnings  
**Location:** `ExcelAPI/ExcelAPI/Rendering/`

---

## Objective

Centralize all Excel style interpretation before ImageEngine. Replace inline style parsing (previously scattered across OpenXmlParser, FillEngine, BorderEngine, TextEngine) with a single StyleResolver system that consumes `ResolvedCellStyle` — never OpenXml structures directly.

---

## Architecture

**Before:**
```
OpenXmlParser (ApplyCellStyle inline + duplicate ResolveColor)
  ↓
FillEngine    parses fills independently
BorderEngine  parses borders independently
TextEngine    parses fonts independently
  ↓
Duplicated style logic across all layers
```

**After:**
```
OpenXmlParser → StyleResolver.Resolve() + StyleCache
  ↓
ResolvedCellStyle (canonical, cached per CellFormat index)
  ↓
FillEngine    reads FillColorArgb, PatternType from ResolvedCellStyle
BorderEngine  reads Border from ResolvedCellStyle
TextEngine    reads Font from ResolvedCellStyle
  ↓
Single authoritative style resolution
```

---

## Files Created (5 new)

| File | Purpose |
|:-----|:--------|
| `Rendering/ResolvedCellStyle.cs` | Canonical resolved style model — Font, Fill, Border, Alignment. All render layers consume this. |
| `Rendering/ColorResolver.cs` | Resolves RGB/Indexed/Theme/Auto colors to #AARRGGBB. Delegates theme resolution to ThemeResolver. |
| `Rendering/ThemeResolver.cs` | Reads theme1.xml, resolves 12 theme colors (Dark1/2, Light1/2, Accent1-6, Hyperlink, FollowedHyperlink) with tint/shade using ECMA-376 algorithm. **Solves the 11C limitation.** |
| `Rendering/StyleCache.cs` | Thread-safe `Dictionary<int, ResolvedCellStyle>`. Each CellFormat resolved once. |
| `Rendering/StyleResolver.cs` | Orchestrates all style resolution from OpenXml Stylesheet. `PreResolveAll()` for eager caching (all styles resolved at parse time). |

## Files Modified (3 changed)

| File | Change |
|:-----|:-------|
| `Rendering/WorkbookModel.cs` | `RenderCell` now has `ResolvedStyle` property (`ResolvedCellStyle?`) |
| `Rendering/OpenXmlParser.cs` | Replaced old `ApplyCellStyle` (200+ lines) with `StyleResolver.Resolve()`. Removed dead code (`ParseBorderItem`, `ResolveColor`, `IndexedColorToArgb`, `ComputeGeometry`). Added `StyleResolver` constructor parameter. Calls `LoadTheme()` at workbook level. |
| `Program.cs` | Registered `ThemeResolver`, `ColorResolver`, `StyleCache`, `StyleResolver` |

---

## Theme Support Matrix

| Theme Color | Index | Support | Sample Default |
|:------------|:-----:|:-------:|:---------------|
| Dark1 | 0 | ✅ | `#FF000000` |
| Light1 | 1 | ✅ | `#FFFFFFFF` |
| Dark2 | 2 | ✅ | `#FF44546A` |
| Light2 | 3 | ✅ | `#FFF2F2F2` |
| Accent1 | 4 | ✅ | `#FF4472C4` |
| Accent2 | 5 | ✅ | `#FFED7D31` |
| Accent3 | 6 | ✅ | `#FFA5A5A5` |
| Accent4 | 7 | ✅ | `#FFFFC000` |
| Accent5 | 8 | ✅ | `#FF5B9BD5` |
| Accent6 | 9 | ✅ | `#FF70AD47` |
| Hyperlink | 10 | ✅ | `#FF0563C1` |
| FollowedHyperlink | 11 | ✅ | `#FF954F72` |
| Tint/Shade | - | ✅ | ECMA-376: blend toward white (+) or black (-) |

---

## ResolvedCellStyle Properties

| Category | Properties |
|:---------|:-----------|
| Font | `FontName`, `FontSize`, `Bold`, `Italic`, `Underline`, `Strikeout`, `FontColorArgb` |
| Fill | `FillColorArgb`, `PatternType` |
| Border | `ResolvedBorder` (Left, Right, Top, Bottom — each with Style + ColorArgb) |
| Alignment | `HorizontalAlignment`, `VerticalAlignment`, `WrapText`, `Indent`, `TextRotation` |

---

## Validation

| Metric | Result |
|:-------|:------:|
| Build errors | **0** ✅ |
| Legacy ApplyCellStyle removed | ✅ — replaced by `StyleResolver.Resolve()` |
| Dead code removed | ✅ — `ParseBorderItem`, `ResolveColor`, `IndexedColorToArgb`, `ComputeGeometry` |
| Theme colors now resolve | ✅ — `LoadTheme(WorkbookPart)` called per workbook |
| Color resolution priority | ✅ — RGB → Theme → Indexed → Auto → null |
| Style caching | ✅ — `StyleCache.GetOrAdd()` resolves each CellFormat once |

---

## Remaining Limitations (before ImageEngine)

| Issue | Notes |
|:------|:------|
| ShrinkToFit not in ResolvedCellStyle | Not yet resolved from alignment |
| Font Charset/Family/Scheme | Not stored — add if ImageEngine needs it |
| Number formats | Not yet prepared — add NumberFormatId for Date/Currency |
| Render layers still read individual properties | Future refactoring: consume `ResolvedStyle` directly |
