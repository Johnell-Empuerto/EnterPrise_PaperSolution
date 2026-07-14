# Diagnostic Report

**Status:** Issues Found
**Severity:** high

## Summary
At least one critical issue detected. Coordinate accuracy may be compromised.

## Issues

| Category | Severity | Detail | Fields Affected |
|----------|----------|--------|-----------------|
| Workbook geometry mismatch | high | No explicit column widths in workbook XML. Engine uses default width which may not match Excel's ... | A1, C1, E1, F1, G1, E2, F2, A3, E3, F3, E4, F4, A5, B5, C... |
| Workbook geometry mismatch | medium | Some columns use default width. Excel may render at different width than our calculation. | A1, C1, E1, F1, G1, E2, F2, A3, E3, F3, E4, F4, A5, B5, C... |

## Recommendations

- Ensure workbook has explicit column widths for all print area columns