# Phase 5.5.3 — Forensic Investigation: Excel Repair Root Cause Identified

**Date:** 2026-07-20  
**Status:** Root Cause Identified — Fix Ready  
**Evidence:** Three-way OOXML forensic comparison (Original vs ConMas vs Generated)

---

## Root Cause

**The rId relationship ID in workbook.xml does NOT match the rId in workbook.xml.rels for the ExcelOutputSetting sheet.**

This is the EXACT cause of Excel's "We found a problem with some content..." repair dialog.

### The Mismatch

When a user opens the workbook, Excel reads:

**`xl/workbook.xml`** (says):
```
Sheet name=ExcelOutputSetting → id=rId3
```

**`xl/_rels/workbook.xml.rels`** (says):
```
rId3 → Target=theme/theme1.xml   (NOT a worksheet!)
rId6 → Target=worksheets/sheet3.xml
```

Excel looks up rId3 for the ExcelOutputSetting sheet, finds theme/theme1.xml (not a worksheet part), and enters repair mode.

### Evidence

Confirmed by direct ZIP inspection of `FormTest - Copy_edited (4).xlsx`:

**Generated workbook.xml.rels:**
```
rId1 → worksheets/sheet1.xml
rId2 → worksheets/sheet2.xml
rId3 → theme/theme1.xml              ← MISMATCH: rId3 NOT sheet3
rId4 → styles.xml
rId5 → sharedStrings.xml
rId6 → worksheets/sheet3.xml          ← sheet3 is here at rId6
```

**Generated workbook.xml SHEETS:**
```
name=_Fields sheetId=2 id=rId1      ← correct
name=Sheet1 sheetId=1 id=rId2       ← correct
name=ExcelOutputSetting sheetId=3 id=rId3  ← WRONG! rId3 points to theme, not sheet3
```

### Why This Happens (The Bug)

Two methods in `PostProcessZipForConMas` compute the new relationship ID from **different sources**:

| Method | Computes new rId from | Result |
|--------|----------------------|--------|
| `UpdateWorkbookXmlForSheet3()` | Sheet elements in workbook.xml (maxRelId = 2) | New relId = rId3 |
| `UpdateWorkbookRelsForSheet3()` | Relationship elements in rels file (maxRelId = 5) | New relId = rId6 |

Both methods must use the **same source** — the rels file (`xl/_rels/workbook.xml.rels`) — which contains ALL relationships including theme, styles, and sharedStrings.

---

## How ConMas Does It Differently

The legacy ConMas workbook was built from scratch by COM Interop, which assigns relationships in a specific order:

**ConMas workbook.xml.rels:**
```
rId1 → worksheets/sheet1.xml     (Worksheet)
rId2 → worksheets/sheet2.xml     (Worksheet)
rId3 → worksheets/sheet3.xml     (ExcelOutputSetting ← CORRECT)
rId4 → theme/theme1.xml          (Theme — after sheets)
rId5 → styles.xml
rId6 → sharedStrings.xml
```

**ConMas workbook.xml SHEETS:**
```
name=_Fields sheetId=2 id=rId1        ← correct
name=Sheet1 sheetId=1 id=rId2         ← correct
name=ExcelOutputSetting sheetId=3 id=rId3  ← CORRECT! rId3 IS sheet3
```

In the ConMas workbook, worksheets are always rId1, rId2, rId3 because they were added first. Non-worksheet relationships (theme, styles, sharedStrings) get higher rIds.

In our generated workbook, we inherit the original workbook's relationship order where theme happens to be rId3. Our post-processing doesn't account for this.

---

## Fix

### The exact fix in `UpdateWorkbookXmlForSheet3()`:

Replace:
```csharp
// Current (BUG): computes from sheet elements ONLY
var idAttr = (string)sheet.Attribute(ns + "id") ?? 
    (string)sheet.Attribute(rAttr);
if (idAttr != null && idAttr.StartsWith("rId") && int.TryParse(idAttr[3..], out int rid))
    if (rid > maxRelId) maxRelId = rid;
```

With: Read the rels file to find the true max relationship ID.

The simplest correct approach: **Both methods should parse the rels file (`xl/_rels/workbook.xml.rels`) to find the maximum relationship ID.**

### Implementation:

In `UpdateWorkbookXmlForSheet3()`:

```csharp
// Parse xl/_rels/workbook.xml.rels to find max rId
int maxRelId = 0;
var relsEntry = pkg.GetEntry("xl/_rels/workbook.xml.rels");
if (relsEntry != null)
{
    using var r = new StreamReader(relsEntry.Open());
    var relsDoc = XDocument.Parse(r.ReadToEnd());
    XNamespace relNs = relsDoc.Root.Name.Namespace;
    foreach (var rel in relsDoc.Root.Elements(relNs + "Relationship"))
    {
        var idAttr = (string)rel.Attribute("Id");
        if (idAttr != null && idAttr.StartsWith("rId") && 
            int.TryParse(idAttr[3..], out int rid))
            if (rid > maxRelId) maxRelId = rid;
    }
}
string relId = "rId" + (maxRelId + 1);
```

This ensures both methods compute the SAME relId from the SAME source.

---

## Verification

After the fix, the generated workbook should have:

**workbook.xml.rels:**
```
rId1 → worksheets/sheet1.xml
rId2 → worksheets/sheet2.xml
rId3 → theme/theme1.xml
rId4 → styles.xml
rId5 → sharedStrings.xml
rId6 → worksheets/sheet3.xml    ← added consistently
```

**workbook.xml SHEETS:**
```
name=_Fields sheetId=2 id=rId1
name=Sheet1 sheetId=1 id=rId2
name=ExcelOutputSetting sheetId=3 id=rId6  ← matches rels!
```

Excel will look up rId6, find worksheets/sheet3.xml, and open the workbook without repair.

---

## Other Differences (Not Root Cause)

The three-way comparison also found these differences between ConMas and Generated, but they are **NOT** the cause of Excel repair:

| Difference | Severity | Notes |
|-----------|----------|-------|
| Shared string count (43 vs 7) | **Not a repair cause** | ConMas rebuilds all strings; our pipeline uses original shared strings |
| Styles (8 borders vs 2) | **Not a repair cause** | Our ZIP restore preserves original styles — correct behavior |
| Comment UIDs | **Not a repair cause** | Unique per generation; Excel ignores UID differences |
| VML positioning | **Not a repair cause** | Minor position differences in comment shapes |
| docProps timestamps | **Not a repair cause** | Metadata fields; Excel does not repair for timestamp differences |

The rId mismatch is the ONLY issue that causes Excel to repair the workbook.

---

## Summary

| Item | Status |
|------|--------|
| Root cause identified | ✅ rId mismatch between workbook.xml and workbook.xml.rels |
| Evidence | ✅ Direct ZIP inspection of both workbooks |
| Fix location | `WorkbookValueWriter.cs` — `UpdateWorkbookXmlForSheet3()` method |
| Fix scope | Change rId computation source from sheet elements to rels file |
| Impact of fix | ExcelOutputSetting sheet uses correct rId (rId6) matching the rels file |
| Expected result | Excel opens without repair dialog |
