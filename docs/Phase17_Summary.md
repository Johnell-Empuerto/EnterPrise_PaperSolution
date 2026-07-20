# Phase 17 — Functional Round-Trip Compatibility (ConMas Behavior)

**Date:** July 20, 2026
**Status:** Specification Complete — Reinforces Phase 16 Direction
**Code Changes:** None (specification/requirements phase)

---

## Objective

Eliminate all remaining assumptions that workbook XML must remain structurally identical. The goal is no longer byte-level fidelity — it is functional round-trip compatibility matching ConMas Designer behavior.

---

## New Success Criteria

PaperLess is considered successful if ALL of these are true:

### 1. Workbook Always Opens in Excel ✓

No corruption. No repair dialog. No warnings.

### 2. Background Remains Identical ✓

Every generation preserves:
- Print Area
- Margins
- Page Size
- Orientation
- Scaling
- Hidden rows
- Hidden columns
- Merged cells
- Images
- Shapes

Pixel perfect across unlimited generations.

### 3. Yellow Editable Fields Remain Aligned ✓

Every export preserves:
- Left
- Top
- Width
- Height
- Cell mapping

No drift across unlimited generations.

### 4. Configuration Survives Forever ✓

Must preserve:
- `_Fields`
- `ExcelOutputSetting`
- Comments
- VML
- Defined Names

Never recreate. Never duplicate. Never lose.

### 5. Workbook is Re-uploadable ✓

The exported workbook must be accepted by PaperLess again and again, without modification.

### 6. No Structural Comparison ✓

**DELETE the assumption that these must match:**

| ❌ Stop Checking | ✅ Start Checking |
|-----------------|------------------|
| XML order | Can Excel open it? |
| Relationship IDs | Can PaperLess read it? |
| Shared string order | Can fields be reconstructed? |
| Style IDs | Can background be rendered? |
| ZIP ordering | Can it export again? |
| Font IDs | Is configuration preserved? |
| SHA hashes | — |

These are implementation details. Modern Office regenerates many of them naturally. ConMas itself does not validate them.

---

## New Compatibility Validator Requirements

The validator should only verify critical business requirements:

| Check | Required | Status |
|-------|----------|--------|
| Workbook opens successfully | ✅ | Implemented in Phase 16 |
| At least one printable worksheet exists | ✅ | Implemented in Phase 16 |
| `_Fields` exists | ✅ | Implemented in Phase 16 |
| `ExcelOutputSetting` exists | ✅ | Implemented in Phase 16 |
| Comments preserved OR `_Field` definitions preserved | ✅ | Implemented in Phase 16 |
| PrintArea exists | ⚠️ | Partially — PageSetup checked, PrintArea not yet |
| PageSetup exists | ✅ | Implemented in Phase 16 |
| Background PNG renders | 🔲 | Future enhancement |
| Field count preserved | 🔲 | Future enhancement |
| Workbook saves | ✅ | Implicit — write succeeds |
| Workbook can be uploaded again | 🔲 | Acceptance test needed |

---

## New Acceptance Test

The project is only considered complete when this passes:

```
Generation 0: Original Template
  ↓ Export
Generation 1: Export1.xlsx
  ↓ Upload → Export
Generation 2: Export2.xlsx
  ↓ Upload → Export
Generation 3: Export3.xlsx
  ↓ ...
  ↓
Generation 50: Export50.xlsx
```

Every generation must satisfy:
- ✅ Background identical
- ✅ Fields identical
- ✅ Configuration preserved
- ✅ No corruption
- ✅ Upload succeeds
- ✅ Export succeeds

---

## Stop Chasing XML

Do not spend time trying to make these identical:

- `workbook.xml`
- `styles.xml`
- `sharedStrings.xml`
- relationship IDs
- ZIP ordering

**Unless there is concrete evidence that one of them actually breaks import.**

Every hour spent cloning OpenXML internals is an hour not spent improving the actual workflow.

---

## Focus on Business Compatibility

From now on, PaperLess should validate only what the business actually needs:

| Question | Priority |
|----------|----------|
| Can the workbook be edited? | 🔴 Critical |
| Can it be rendered? | 🔴 Critical |
| Can it be exported? | 🔴 Critical |
| Can it be uploaded again? | 🔴 Critical |
| Are all form definitions preserved? | 🔴 Critical |
| Does the user experience match ConMas? | 🟡 Important |

If the answer is **yes**, then the workbook is compatible — even if its internal XML differs from the original.

---

## Alignment with Phase 16 Implementation

| Phase 16 Implementation | Phase 17 Requirement | Status |
|------------------------|---------------------|--------|
| `CompatibilityValidator` created | Replace structural validation | ✅ Done |
| Workbook opens check | Workbook opens successfully | ✅ Done |
| Printable sheets check | At least one printable sheet | ✅ Done |
| Config sheets check | `_Fields` and `ExcelOutputSetting` | ✅ Done |
| Field metadata check | Comments or _Field definitions | ✅ Done |
| Page setup check | PrintArea and PageSetup | ⚠️ Partial |
| Never throws | No rejection on structural difference | ✅ Done |
| Always returns workbook | Workbook always downloadable | ✅ Done |
| Field count preservation | Every field must still exist | 🔲 Not yet implemented |
| Background alignment | Pixel-perfect across generations | 🔲 Not yet implemented |
| Multi-gen acceptance test | 50-generation test | 🔲 Not yet implemented |

---

## Next Steps (Recommended)

1. **Add field count check to `CompatibilityValidator`** — Verify that the number of comments/fields in the edited workbook matches the original
2. **Add print area/dimension validation** — Verify that page size, margins, print area are preserved
3. **Run multi-generation acceptance test** — Export → Upload → Export 50 times and verify all success criteria
