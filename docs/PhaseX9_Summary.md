# Phase X.9 Summary — Forensic OOXML Binary Diff & Dependency Analysis

**Date:** 2026-07-17
**Type:** Pure investigation — zero code changes
**Goal:** Determine why the legacy ConMas Output Excel workbook round-trips perfectly while our generated workbook reconstructs only 1 field

---

## What Was Done

### 1. Created `forensic_diff.py` (Standalone Analysis Tool)
A Python script that:
- Opens both legacy (`FormTest - Copy.xlsx`) and generated (`FormTest - Copy_output.xlsx`) workbooks as ZIP files
- Compares every ZIP entry by size and content
- Performs XML node-level comparison using ElementTree
- Compares worksheet structures (dimensions, merges, rows, columns)
- Compares comments (count, authors, line structure, address locations)
- Compares defined names, styles, shared strings
- Produces a dependency graph and key differences summary

### 2. Key Findings

| Finding | Severity |
|---------|----------|
| Legacy has **3 sheets** (Sheet1, _Fields, ExcelOutputSetting); generated has **2** (missing ExcelOutputSetting) | HIGH |
| Legacy Sheet1 has **9 rows, 33 cells, 5 merged regions**; generated Sheet1 has **1 row, 0 cells, 0 merges** | **CRITICAL** |
| Legacy has **6 comments** (one per field); generated had only **1** (A12) before the Phase X.5 fix | PREVIOUSLY CRITICAL |
| Legacy comment format: **~25 lines** with name, type, cluster index, reserved fields, parameters; ours: **4 lines** with only name, type, placeholder | HIGH |
| Legacy has **PrintArea** defined name (`Sheet1!$A$1:$D$12`); generated has **none** | **CRITICAL** |
| **_Fields sheet**: Legacy has headers only (1 row); generated has 7 rows (header + 6 data) | LOW — ConMas skips hidden sheets |
| **ExcelOutputSetting**: Legacy has 36 rows with XML config fragments; generated has **none** | UNKNOWN — needs runtime testing |

### 3. Root Cause Chain

```
WorkbookGenerator creates Sheet1 with no merged cells, no PrintArea, no cell values
    ↓
Comments written to multi-cell ranges (before Phase X.5 fix) → 5 of 6 fail
    ↓
Only cell A12 (single cell) gets a comment → only 1 cluster reconstructed
    ↓
Even with Phase X.5 fix (6 comments): no merged cells, no PrintArea → ConMas loader still fails
```

### 4. Metadata Dependency Graph (Evidence-Based)

```
WORKBOOK
├── xl/workbook.xml
│   ├── Sheet1 (visible) ─── REQUIRED
│   ├── _Fields (hidden) ─── IGNORED by ConMas (skips hidden sheets)
│   └── ExcelOutputSetting ── UNKNOWN (needs runtime test)
├── xl/worksheets/sheet1.xml
│   ├── mergeCells ─── REQUIRED (5 regions define field boundaries)
│   ├── sheetData ─── OPTIONAL (existing values)
│   └── printArea ─── REQUIRED (processing boundary)
├── xl/comments1.xml ─── PRIMARY metadata source (REQUIRED)
│   ├── Line 0: Field Name ─── REQUIRED
│   ├── Line 1: Field Type ─── REQUIRED
│   ├── Line 2: Cluster Index ─── LIKELY REQUIRED
│   ├── Lines 3-4: Reserved ─── OPTIONAL
│   └── Line 5+: Parameters ─── REQUIRED for full config
└── DefinedNames
    └── _xlnm.Print_Area ─── REQUIRED
```

### 5. Files Created

| File | Purpose |
|------|---------|
| `forensic_diff.py` | Standalone OOXML comparison tool (one-time use) |
| `docs/PhaseX9_Forensic_Diff_Dependency_Report.md` | Full forensic report with implementation roadmap |
| `docs/PhaseX9_Summary.md` | This file |

### 6. What Was NOT Done (By Design)
- No changes to `WorkbookGenerator.cs`
- No changes to `WorkbookReaderService.cs`
- No changes to any API endpoint
- No changes to the frontend
- No implementation of ExcelOutputSetting
- No comment format changes

The phase was pure investigation to build the evidence base for a precise implementation in the next phase.

---

## Next Step Recommended

The investigation is complete. The next phase should implement the **critical fixes** identified in the dependency graph:
1. Add merged cells to Sheet1
2. Add PrintArea to Sheet1
3. Add cell values to Sheet1
4. Add cluster index and parameters to comment format
5. Add ExcelOutputSetting sheet
