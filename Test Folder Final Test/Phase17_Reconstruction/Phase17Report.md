# Phase 17 — Legacy Coordinate Reconstruction Report

**Algorithm:** Column-Width/Row-Height Summation (from decompiled Cimtops.Excel.dll)
**Date:** 2026-07-12 20:43:18

## Formula

```
left_pt   = Σ Col[n].Width  for n < cluster.StartColumn
top_pt    = Σ Row[n].Height for n < cluster.StartRow
width_pt  = Σ Col[n].Width  for n = StartColumn..EndColumn
height_pt = Σ Row[n].Height for n = StartRow..EndRow
ratio     = RoundEx((pos_pt + origin_pt) / page_dim_pt, 7)
```

**Key differences from Phase 10:**
- Position = sum of column widths / row heights (NOT Range.Left/Range.Top)
- Width/Height = sum of column widths / row heights for the spanned range
- Normalization divisor = PageSetup.PageWidth/PageHeight

---

## Per-Template Detail

### Template 546

| Cluster | DB L | DB T | DB R | DB B | Gen L | Gen T | Gen R | Gen B | L? | T? | R? | B? |
|---------|------|------|------|------|-------|-------|-------|-------|:--:|:--:|:--:|:--:|
| $A$1:$B$2 | 0.3364706 | 0.3845454 | 0.4982353 | 0.4218182 | 0.3364706 | 0.3845454 | 0.4933333 | 0.4216666 | ✅ | ✅ | ❌ | ❌ |
| $C$1:$D$2 | 0.5000000 | 0.3845454 | 0.6635294 | 0.4218182 | 0.4933333 | 0.3845454 | 0.6501961 | 0.4216666 | ❌ | ✅ | ❌ | ❌ |
| $A$3:$D$4 | 0.3364706 | 0.4231818 | 0.6635294 | 0.4604546 | 0.3364706 | 0.4216666 | 0.6501961 | 0.4595454 | ✅ | ❌ | ❌ | ❌ |
| $A$6:$D$7 | 0.3364706 | 0.4809091 | 0.6635294 | 0.5181818 | 0.3364706 | 0.4784848 | 0.6501961 | 0.5163636 | ✅ | ❌ | ❌ | ❌ |
| $A$9:$D$10 | 0.3364706 | 0.5386364 | 0.6635294 | 0.5759091 | 0.3364706 | 0.5353030 | 0.6501961 | 0.5731817 | ✅ | ❌ | ❌ | ❌ |
| $A$12 | 0.3364706 | 0.5963637 | 0.4164706 | 0.6150000 | 0.3364706 | 0.5921212 | 0.4149020 | 0.6110606 | ✅ | ❌ | ❌ | ❌ |

**Results:** Left=5/6 (83%), Top=2/6 (33%), Right=0/6 (0%), Bottom=0/6 (0%)

---

### Template 547

| Cluster | DB L | DB T | DB R | DB B | Gen L | Gen T | Gen R | Gen B | L? | T? | R? | B? |
|---------|------|------|------|------|-------|-------|-------|-------|:--:|:--:|:--:|:--:|
| $I$6:$M$6 | 0.5294118 | 0.1654546 | 0.9282353 | 0.1868182 | 0.5294118 | 0.1654546 | 0.9215687 | 0.1843940 | ✅ | ✅ | ❌ | ❌ |
| $I$7:$M$7 | 0.5294118 | 0.1877273 | 0.9282353 | 0.2095454 | 0.5294118 | 0.1843940 | 0.9215687 | 0.2033334 | ✅ | ❌ | ❌ | ❌ |
| $I$8:$M$8 | 0.5294118 | 0.2104545 | 0.9282353 | 0.2327273 | 0.5294118 | 0.2033334 | 0.9215687 | 0.2222728 | ✅ | ❌ | ❌ | ❌ |
| $I$9:$M$9 | 0.5294118 | 0.2336364 | 0.9282353 | 0.2559091 | 0.5294118 | 0.2222728 | 0.9215687 | 0.2412122 | ✅ | ❌ | ❌ | ❌ |
| $I$10:$M$10 | 0.5294118 | 0.2568182 | 0.9282353 | 0.2786364 | 0.5294118 | 0.2412122 | 0.9215687 | 0.2601516 | ✅ | ❌ | ❌ | ❌ |

**Results:** Left=5/5 (100%), Top=1/5 (20%), Right=0/5 (0%), Bottom=0/5 (0%)

---

### Template 548

| Cluster | DB L | DB T | DB R | DB B | Gen L | Gen T | Gen R | Gen B | L? | T? | R? | B? |
|---------|------|------|------|------|-------|-------|-------|-------|:--:|:--:|:--:|:--:|
| $A$10:$D$11 | 0.0847059 | 0.2040909 | 0.4111765 | 0.2418182 | 0.0847059 | 0.2040909 | 0.3984314 | 0.2404545 | ✅ | ✅ | ❌ | ❌ |
| $E$10:$G$11 | 0.4129412 | 0.2040909 | 0.6582353 | 0.2418182 | 0.3984314 | 0.2040909 | 0.6337255 | 0.2404545 | ❌ | ✅ | ❌ | ❌ |

**Results:** Left=1/2 (50%), Top=2/2 (100%), Right=0/2 (0%), Bottom=0/2 (0%)

---

## Cross-Template Summary

| Template | Left% | Top% | Right% | Bottom% | Overall |
|----------|:----:|:----:|:------:|:-------:|:-------:|
| 546 | 83.3% | 33.3% | 0.0% | 0.0% | 0.0% |
| 547 | 100.0% | 20.0% | 0.0% | 0.0% | 0.0% |
| 548 | 50.0% | 100.0% | 0.0% | 0.0% | 0.0% |

## Conclusion

The column-width/row-height summation algorithm was implemented and tested against
all three reference templates. See per-template detail above for exact match results.

**If Left/Top/Right/Bottom do not all match at 100%:**

1. Check if the PrintArea boundary was correctly identified
2. Check if column widths and row heights from COM match OpenXML values
3. Check if the printed origin offset (page margins + centering) is correct
4. The publishing code in ConMasClient.exe may apply additional transforms
