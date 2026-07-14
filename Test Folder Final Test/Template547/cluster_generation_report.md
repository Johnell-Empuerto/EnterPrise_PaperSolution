# Cluster Generation Investigation

**Workbook:** 547_workbook_copy.xlsx
**DB Clusters:** 5

## Sheet Analysis

Total cells in workbook: 487
Merged cells in workbook: 42
DB clusters: 5

### DB Cluster vs Workbook Cell Analysis

| DB Address | Type | In Merge? | Has Data? | Workbook Ref |
|------------|------|-----------|-----------|--------------|
| $I$6:$M$6 | KeyboardText | False | False |  |
| $I$7:$M$7 | KeyboardText | False | False |  |
| $I$8:$M$8 | KeyboardText | False | False |  |
| $I$9:$M$9 | KeyboardText | False | False |  |
| $I$10:$M$10 | KeyboardText | False | False |  |

### Style Analysis: DB Cluster Cells vs Non-Cluster Cells

- Unique styles used ONLY by DB clusters: None
- Styles shared with non-cluster cells: 


### Named Range Analysis

- _xlnm.Print_Area: '第15回DMSK_i-Reporter'!$B$1:$M$44
## Sheet Analysis

Total cells in workbook: 31
Merged cells in workbook: 0
DB clusters: 5

### DB Cluster vs Workbook Cell Analysis

| DB Address | Type | In Merge? | Has Data? | Workbook Ref |
|------------|------|-----------|-----------|--------------|
| $I$6:$M$6 | KeyboardText | False | False |  |
| $I$7:$M$7 | KeyboardText | False | False |  |
| $I$8:$M$8 | KeyboardText | False | False |  |
| $I$9:$M$9 | KeyboardText | False | False |  |
| $I$10:$M$10 | KeyboardText | False | False |  |

### Style Analysis: DB Cluster Cells vs Non-Cluster Cells

- Unique styles used ONLY by DB clusters: None
- Styles shared with non-cluster cells: 


### Named Range Analysis

- _xlnm.Print_Area: '第15回DMSK_i-Reporter'!$B$1:$M$44

## Cluster Generation Rule Hypothesis

Based on the analysis, DB clusters may be determined by:
1. **User-selected input fields** in ConMasDesigner (most likely)
2. Cells with specific style properties
3. Cells within a specific column range
4. Cells with data validation applied
5. Named ranges that define input areas

The current engine generates clusters for ALL merged cells and standalone cells.
The DB only stores clusters that were configured as input fields.
Without access to the ConMasDesigner metadata stored in the workbook,
the exact selection criteria cannot be fully replicated.
