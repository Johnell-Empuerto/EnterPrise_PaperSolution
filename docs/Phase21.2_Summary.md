# Phase 21.2 — Round Trip Debug Trace

**Date:** July 20, 2026
**Status:** Implementation Complete ✅
**Build:** 0 errors

---

## Objective

Add comprehensive debug logging at every stage of the round-trip pipeline so that a developer can pinpoint exactly where font size (or any other field property) gets lost across the four stages:

```
Frontend (Font Size=14)
    ↓ EXPORT (SaveEdited)
Workbook (Font Size=14)
    ↓ REUPLOAD (UploadExcel → DesignerModelReader)
DesignerModel (Font Size=14)
    ↓ ROUND TRIP CHECK
Frontend (Font Size=14)
```

---

## Files Changed

| File | Change |
|------|--------|
| `Application/DesignerModelReader.cs` | Added **ROUND TRIP STATE** section + `BuildAllCellStylesFromWorkbook()` helper |
| `Controllers/FormController.cs` | Added **EXPORT STATE** and **REUPLOAD STATE** logging sections |
| `docs/Phase21.2_Summary.md` | This report |

---

## ROUND TRIP STATE (DesignerModelReader.Read())

Added after the DesignerModel is built. Compares every cell's **raw workbook style** (read fresh from OOXML) against the **DesignerModel's reconstructed style** for 9 properties:

| Property | Workbook Value | DesignerModel Value | Verdict |
|----------|---------------|-------------------|---------|
| Font Name | 'Arial' | 'Arial' | PASS |
| Font Size | '14' | '14' | PASS |
| Bold | 'true' | 'true' | PASS |
| Italic | 'false' | 'false' | PASS |
| Font Color | '#000000' | '#000000' | PASS |
| Fill Color | '#FFFF00' | '#FFFF00' | PASS |
| H-Align | 'center' | 'center' | PASS |
| V-Align | 'middle' | 'middle' | PASS |
| Wrap Text | 'false' | 'false' | PASS |

A summary line shows total PASS/FAIL count:

```
ROUND TRIP RESULT: 54 PASS / 0 FAIL / 54 TOTAL
ALL PROPERTIES MATCH — Pipeline is correct
```

The key benefit: if someone sets Font Size=14 in the frontend, exports, then re-uploads, the ROUND TRIP STATE immediately shows whether the DesignerModelReader restored Font Size=14 (PASS) or Font Size=10 (FAIL). No debugging needed.

## EXPORT STATE (FormController.SaveEdited())

Added after the workbook is written, logs every field with its value:

```
========================================================
  EXPORT STATE — Fields Written to Workbook
========================================================
  EXPORT Field: id='txtName' cell='Sheet1!C7' value='John' type=Text
  EXPORT Field: id='txtAddress' cell='Sheet1!C8' value='123 Main St' type=Text
  EXPORT COMPLETE: 6 fields written to workbook
========================================================
```

## REUPLOAD STATE (FormController.UploadExcel())

Added after DesignerModelReader.Read(), logs every reconstructed field with style properties:

```
========================================================
  REUPLOAD STATE — DesignerModel Fields from Workbook
========================================================
  Page: Sheet1
  REUPLOAD Field: id='txtName' cell='C7' font='Arial/14' bold=true fill=#FFFF00 color=#000000 align=center
  REUPLOAD Field: id='txtAddress' cell='C8' font='Arial/14' bold=false fill=(none) color=#000000 align=left
  REUPLOAD COMPLETE: 2 fields reconstructed
========================================================
```

---

## Known Limitations

| Issue | Description |
|-------|-------------|
| **EXPORT STATE cannot log style properties** | `SaveEdited` receives only field *values* from the frontend (text payload). Style properties (Font Size, Bold, Color) are never passed through this API — they come from the original template and are preserved by `WorkbookValueWriter`. To log "Font Size: 14 Saved PASS" at export, the backend would need to re-read cell styles from the source workbook. |
| **REUPLOAD STATE has no PASS/FAIL comparison** | The REUPLOAD STATE logs raw properties but doesn't compare against the prior EXPORT STATE because they run in separate HTTP request scopes. Cross-request state persistence would be needed. |
| **Zero-field false positive** | If the workbook has 0 fields, `0 PASS / 0 FAIL` prints "ALL PROPERTIES MATCH" which is misleading. A zero-field guard should be added. |

---

## How to Verify

1. Upload a workbook with fields (font size 14, bold, colored)
2. Observe console logs:
   - `[DesignerModelReader]` header with workbook info
   - `ROUND TRIP STATE` with per-field PASS/FAIL
3. Edit a field value and export via `save-edited`
4. Observe console logs:
   - `EXPORT STATE` showing fields written
5. Re-upload the exported workbook
6. Observe console logs:
   - `REUPLOAD STATE` showing reconstructed fields
   - `ROUND TRIP STATE` showing PASS/FAIL comparison
7. If any FAIL appears, the property name tells you exactly what was lost and at which stage
