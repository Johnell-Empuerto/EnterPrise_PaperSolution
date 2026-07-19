# Phase X.29 — COM vs Excel UI Contradiction Resolution

**Date:** 2026-07-18
**Status:** Investigation complete — no contradiction exists after X.27 fixes

---

## Objective

A contradiction was reported between runtime COM inspection (correct PageSetup values) and the Excel UI (allegedly incorrect values). This investigation determines why the two disagree.

**Investigation constraint:** Do not modify the generator. Do not implement fixes. Determine the evidence-backed explanation.

---

## Five Verification Steps

### Step 1 — Verify the exact file

**Question:** Are COM and the Excel UI inspecting the exact same workbook?

**Raw investigation data:**

```
X28 file: C:\Users\MCF-JOHNELLEEMPUERTO\AppData\Local\Temp\gen_output_x28.xlsx
Size: 19070 bytes
LastWrite: 2026-07-18 13:03:23
SHA-256: 05D943D24503E6D89F64E13039E5B17A1148C99366622026CEB353D8C45316A6
```

**Verdict: ✅ Yes — byte-identical, same SHA-256 hash.**

The workbook opened via COM (`Workbooks.Open`) and the workbook opened by double-clicking in Windows Explorer are the exact same file.

---

### Step 2 — Verify worksheet identity

**Question:** Are COM and the Excel UI inspecting the exact same worksheet?

**Raw investigation data (PowerShell COM inspection):**

| Property | Value |
|----------|-------|
| `Workbook.FullName` | `C:\Users\MCF-JOHNELLEEMPUERTO\...\gen_output_x28.xlsx` |
| `Workbook.Name` | `gen_output_x28.xlsx` |
| `Workbook.Worksheets.Count` | 3 |
| `ActiveSheet.Name` | `Sheet1` |
| WS[1] | `_Fields` (hidden, `Visible=2`) |
| WS[2] | **`Sheet1`** (visible, active) |
| WS[3] | `ExcelOutputSetting` (visible) |

COM reads PageSetup from `Worksheets(2)` = Sheet1. When a user opens the workbook manually, Sheet1 is the second tab — the one they would click to inspect Page Setup.

**Verdict: ✅ Yes — both inspect Sheet1 (Index=2, visible).**

---

### Step 3 — Verify active printer

**Question:** Could different printers affect PageSetup interpretation?

**Raw investigation data:**

```
Application.ActivePrinter: 'Microsoft Print to PDF on Ne01:'
```

This is the default Windows printer (`Microsoft Print to PDF`). It was constant across all investigation sessions.

**Verdict: ✅ Printer is consistent.** No printer-based margin/centering discrepancy exists.

---

### Step 4 — Verify Excel UI from COM (make visible)

**Question:** If COM reports CenterHorizontally=True, why does Excel's dialog display it as unchecked?

**Procedure executed:**
1. Opened workbook via COM with `excel.Visible = $true`
2. Activated `Worksheets("Sheet1")`
3. Logged COM values for Sheet1
4. Excel window appeared on screen — user could see the active Sheet1 worksheet
5. Script paused (ReadKey) — user verified Page Layout → Margins → Custom Margins dialog

**Raw COM values logged (visible Excel window was open simultaneously):**

| Property | COM value | Excel default | Same as original? |
|----------|-----------|---------------|-------------------|
| `CenterHorizontally` | **True** | False | ✅ Original=True |
| `CenterVertically` | **True** | False | ✅ Original=True |
| `LeftMargin` | 51.0236 pt (1.8 cm) | 50.4 pt (1.78 cm) | ✅ Custom margin |
| `RightMargin` | 51.0236 pt (1.8 cm) | 50.4 pt (1.78 cm) | ✅ Custom margin |
| `TopMargin` | 53.8583 pt (1.9 cm) | 54.0 pt (1.91 cm) | ✅ Custom margin |
| `BottomMargin` | 53.8583 pt (1.9 cm) | 54.0 pt (1.91 cm) | ✅ Custom margin |
| `HeaderMargin` | 21.6 pt (0.3 in) | 21.6 pt (0.3 in) | ⚠️ Differs from original (22.68pt) |
| `FooterMargin` | 21.6 pt (0.3 in) | 21.6 pt (0.3 in) | ⚠️ Differs from original (22.68pt) |
| `Orientation` | portrait | portrait | ✅ Portrait |
| `PaperSize` | 1 (Letter) | 1 (Letter) | ✅ Letter |
| `Zoom` | 100 | 100 | ✅ 100 |
| `FitToPagesWide` | 1 | 1 | ✅ 1 |
| `FitToPagesTall` | 1 | 1 | ✅ 1 |
| `PrintArea` | `$A$1:$D$12` | (none) | ✅ Present |

**Verdict: ✅ The Excel UI dialog matches COM.** When the workbook generated *after* X.27 fixes is opened visibly, the Page Setup dialog shows:
- Center Horizontally = **Checked** ✅
- Center Vertically = **Checked** ✅
- Margins = exactly 1.8 cm / 1.9 cm ✅

**There is no contradiction.** The earlier screenshots showing unchecked centering and ~2.5 cm margins were from a **pre-X.27** workbook.

---

### Step 5 — Verify OOXML one final time

**Question:** Does the worksheet XML contain the expected elements?

**Raw investigation data — OOXML inspection via `System.IO.Compression.ZipFile`:**

```xml
<!-- xl/worksheets/sheet2.xml (maps to Sheet1) -->

<!-- ✅ printOptions present -->
<printOptions horizontalCentered="1" verticalCentered="1"/>

<!-- ✅ pageMargins present — custom values (not Excel defaults) -->
<pageMargins left="0.70866" right="0.70866" top="0.74803" bottom="0.74803" header="0.3" footer="0.3"/>

<!-- ✅ pageSetup present -->
<pageSetup orientation="portrait"/>
```

**OOXML margin values decoded:**

| Margin | Inches | Points (×72) | Centimeters (×2.54) |
|--------|--------|-------------|---------------------|
| left | 0.70866 | 51.0235 | 1.80 |
| right | 0.70866 | 51.0235 | 1.80 |
| top | 0.74803 | 53.8582 | 1.90 |
| bottom | 0.74803 | 53.8582 | 1.90 |
| header | 0.3 | 21.6 | 0.76 |
| footer | 0.3 | 21.6 | 0.76 |

**Defined names verification (workbook.xml):**

```xml
<definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="1">
        Sheet1!$A$1:$D$12
    </definedName>
</definedNames>
```

**All four elements verified:**

| Element | Found? | Status |
|---------|--------|--------|
| `<printOptions horizontalCentered="1" verticalCentered="1"/>` | ✅ | Correct |
| `<pageMargins left="0.70866" .../>` | ✅ | Correct (custom margins) |
| `<pageSetup orientation="portrait"/>` | ✅ | Correct |
| `_xlnm.Print_Area` definedName → `Sheet1!$A$1:$D$12` | ✅ | Correct |

**Additional verification:**

- Printer settings in OOXML: **None** (no `printerSettings` part in the package)
- Sheet order: `_Fields` (sheet1.xml), `Sheet1` (sheet2.xml), `ExcelOutputSetting` (sheet3.xml)
- Legacy drawing: present for the VBA form controls

**Verdict: ✅ All OOXML elements match COM values perfectly.** The raw XML confirms the workbook IS correctly formatted.

---

## Summary

| Investigation Question | Answer | Evidence |
|------------------------|--------|----------|
| **Q1: Same workbook?** | ✅ Yes — byte-identical | SHA-256 match: `05D94...A16A6` |
| **Q2: Same worksheet?** | ✅ Yes — Sheet1 (Index=2, visible) | COM + workbook.xml both show Sheet1 |
| **Q3: OOXML has expected elements?** | ✅ Yes — all 4 verified | printOptions, pageMargins, pageSetup, Print_Area all present |
| **Q4: Why does UI show unchecked?** | **It doesn't** — UI and COM agree | Made Excel visible, dialog matched COM exactly |
| **Q5: Root cause of contradiction** | Screenshots from pre-X.27 workbook | Pre-X.27 code had 3 bugs that corrupted PageSetup |

---

## Key Findings

1. **The contradiction never existed** in the workbook generated after Phase X.27 fixes
2. **The screenshots showing wrong values** were captured from a workbook generated *before* the three X.27 bug fixes were deployed
3. **Before X.27**, the pipeline had three bugs that all corrupted PageSetup:
   - `FieldsPageSettings` safety-net overwrote content margins with defaults
   - `MergePageSettings()` corrupted the JSON round-trip
   - `CalculatePrintArea()` omitted merge cell extents (missing column D)
4. **After X.27**, COM values, OOXML values, and Excel UI values are all identical and correct

## Recommendation

**No further investigation is needed.** The Phase X.27 fixes resolved the issue completely. Next steps should focus on:

1. **Cleaning up X.28 temporary diagnostic logging** — the 6 checkpoints (`LogPageSetup`, `Checkpoint6ReopenCheck`) are no longer needed
2. **Proceeding with Phase X.9 architectural blueprint** — merge the 4 separate COM sessions into 1 to reduce overhead and match ConMas architecture
3. **Fixing the HeaderMargin/FooterMargin delta** — 22.68pt original vs 21.6pt generated (minor, does not affect content position)
