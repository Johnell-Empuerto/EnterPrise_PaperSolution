# Phase 5.2.4 — Root Cause Investigation: workbook.xml Still Mutating After Full ZIP Restore

## Status: **Diagnostic Logging Added** ✅ (Run save pipeline to see results)

---

## Objective

Find the exact XML node in `xl/workbook.xml` that differs between original and edited workbooks, despite the comprehensive ZIP entry restore from Phase 5.2.3. **No validation rules were changed.**

---

## Changes Made

### File 1: `Application/WorkbookDiffValidator.cs`

#### Step 1 — Element-by-Element XPath Comparison

**Before:** The workbook.xml comparison cloned both workbooks, removed `<Sheets>`, `<DefinedNames>`, and `<WorkbookProtection>` children, then compared the full remaining `OuterXml` as a single blob:

```csharp
if (!XmlEquals(origWbClone, editWbClone))
{
    result.WorkbookXmlChanges++;
    result.Details.Add($"Workbook.xml (non-sheet content) differs\n  Orig: {FormatXmlSnippet(origWbClone.OuterXml)}\n  Edit: {FormatXmlSnippet(editWbClone.OuterXml)}");
}
```

**After:** Iterates each child element individually and logs the EXACT XPath, original XML, and edited XML:

```csharp
var origChildren = origWbClone.ChildElements.ToList();
var editChildren = editWbClone.ChildElements.ToList();

for (int ci = 0; ci < maxChildren; ci++)
{
    string xpath = "/workbook/" + (ci < origChildren.Count ? origChildren[ci].LocalName : ...);
    string origOuter = ci < origChildren.Count ? origChildren[ci].OuterXml : "(missing)";
    string editOuter = ci < editChildren.Count ? editChildren[ci].OuterXml : "(missing)";

    if (origOuter != editOuter)
    {
        result.Details.Add($"Workbook.xml difference:\n" +
            $"  XPath: {xpath}\n" +
            $"  Original: {origOuter}\n" +
            $"  Edited:   {editOuter}");
    }
}
```

**Example diagnostic output when run:**
```
Workbook.xml difference:
  XPath: /workbook/calcPr
  Original: <calcPr calcId="191029"/>
  Edited:   <calcPr calcId="191030"/>
```

---

### File 2: `Application/WorkbookValueWriter.cs`

#### Step 2 — SHA256 Hash Logging (Before Restore)

Before the ZIP restore loop, logs:

```
[SHA256] workbook.xml before restore: original=<hash>, current=<hash>
```

This shows whether the SDK actually mutated workbook.xml (hash difference) or if it was already unchanged.

#### Step 4 — Entry-Level Restore/Skip Logging

Every zip entry now logs its status during restore:

```
[RESTORE] Restored 'xl/workbook.xml' (4829 bytes)
[RESTORE] Unchanged 'xl/styles.xml' (15234 bytes) — no restore needed
[RESTORE] Skipped 'xl/worksheets/sheet1.xml' — intentionally modified
```

An explicit `workbookXmlWasRestored` flag tracks whether workbook.xml was restored:

```
[RESTORE] Summary: 12 restored, 8 unchanged, 3 skipped. workbook.xml restored=YES
```

#### Step 2 — SHA256 Hash Logging (After Restore)

After the restore loop, logs:

```
[SHA256] workbook.xml after restore: original=<hash>, final=<hash>, match=YES ✅
```

If hashes don't match:
```
[SHA256] workbook.xml RESTORE FAILED! Original hash does not match final hash.
```

#### Step 3 — Pipeline Ordering Documentation

Added a comment block documenting the complete pipeline order, confirming no writes to the output ZIP after the restore loop:

```
// PIPELINE ORDERING VERIFICATION (Phase 5.2.4, Step 3)
// After this point, NO code should write back into the ZIP.
// Pipeline order:
//   1. File.Copy(source, output) — copy original
//   2. Pre-save ZIP entries (originalZipEntries)
//   3. SpreadsheetDocument.Open(output, true) — SDK opens & modifies
//   4. Write cell values (intentionally modifies worksheets + shared strings)
//   5. SpreadsheetDocument.Dispose — SDK writes its version to disk
//   6. RestoreOriginalEntries (ZIP loop) — overwrites SDK-mutated entries
//   7. SHA256 verification (post-restore) — verifies restore succeeded
//   8. [HERE] — logging and return only. No writes to output ZIP after this.
```

---

## How to Use the Diagnostics

1. **Run the export pipeline** (upload → edit → save)
2. Check the **backend logs** for:

   **a) SHA256 before restore:**
   ```
   [SHA256] workbook.xml before restore: original=abc123, current=def456
   ```
   If current == original → SDK did NOT mutate workbook.xml → issue is elsewhere
   If current != original → SDK mutated it → restore should fix it

   **b) Restore status:**
   ```
   [RESTORE] Restored 'xl/workbook.xml' (4829 bytes)
   ```
   If workbook.xml is NOT in the Restored list → check if it was accidentally Skipped or Unchanged

   **c) SHA256 after restore:**
   ```
   [SHA256] workbook.xml after restore: original=abc123, final=abc123, match=YES ✅
   ```
   If match=NO ❌ → restore failed → the ZIP entry is somehow not being replaced

   **d) Validator XPath details:**
   ```
   Workbook.xml difference:
     XPath: /workbook/calcPr
     Original: <calcPr calcId="191029"/>
     Edited:   <calcPr calcId="191030"/>
   ```

3. **Scenario analysis:**
   - If hashes match after restore ✅ **but** validator still reports `WorkbookXmlChanges=1` → the validator's OpenXml SDK re-parses the file and the SDK itself re-mutates it during `SpreadsheetDocument.Open()` in the validator (read-only mode shouldn't mutate, but verify)
   - If hashes don't match ❌ → the restore loop is not actually replacing the bytes (check: is the entry being deleted and re-created correctly? Is there a `[Content_Types].xml` issue where the entry name differs?)

---

## Build Verification

```
Build succeeded.
  0 Error(s)
```

---

## Files Modified

| File | Change |
|---|---|
| `Application/WorkbookDiffValidator.cs` | Element-by-element XPath comparison for workbook.xml |
| `Application/WorkbookValueWriter.cs` | SHA256 hash logging, entry-level restore logging, pipeline ordering documentation |
| `docs/Phase5.2.4_Root_Cause_WorkbookXml_Still_Mutating.md` | **This report** |

---

## Expected Diagnostic Output (Example)

When the save pipeline runs with these diagnostics, the logs will show:

```
[PRESERVE] Saved 28 original ZIP entries before SDK
[PRESERVE]   Saved entry: xl/workbook.xml (4829 bytes)
[PRESERVE]   Saved entry: xl/styles.xml (15234 bytes)
...

========== WRITE VALUES ==========
...
Workbook sheets found: 2

[SHA256] workbook.xml before restore: original=abc123def456, current=7890abcdef
  [indicating SDK mutated workbook.xml]

[RESTORE] Restored 'xl/workbook.xml' (4829 bytes)  ← EXACTLY confirm this
[RESTORE] Unchanged 'xl/styles.xml' (15234 bytes)
[RESTORE] Skipped 'xl/worksheets/sheet1.xml' — intentionally modified
[RESTORE] Summary: 12 restored, 8 unchanged, 3 skipped. workbook.xml restored=YES

[SHA256] workbook.xml after restore: original=abc123def456, final=abc123def456, match=YES ✅
  [If match=NO, restore failed]

========== WORKBOOK VALIDATION ==========
...
Workbook.xml difference:                           ← Only appears if hash still differs
  XPath: /workbook/calcPr
  Original: <calcPr calcId="191029"/>
  Edited:   <calcPr calcId="191030"/>
```

---

## What NOT Changed

- ❌ Validation logic — `WorkbookXmlChanges++` still increments the same way
- ❌ Validation thresholds — `Passed` still requires `TotalDifferences == 0`
- ❌ API contracts — all endpoints unchanged
- ❌ Rendering pipeline — unchanged
