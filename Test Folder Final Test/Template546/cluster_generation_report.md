# Cluster Generation Investigation

**Workbook:** FormTest - Copy.xlsx
**DB Clusters:** 6

## Sheet Analysis

Total cells in workbook: 33
Merged cells in workbook: 5
DB clusters: 6

### DB Cluster vs Workbook Cell Analysis

| DB Address | Type | In Merge? | Has Data? | Workbook Ref |
|------------|------|-----------|-----------|--------------|
| $A$1:$B$2 | KeyboardText | False | False |  |
| $C$1:$D$2 | KeyboardText | False | False |  |
| $A$3:$D$4 | KeyboardText | False | False |  |
| $A$6:$D$7 | KeyboardText | False | False |  |
| $A$9:$D$10 | KeyboardText | False | False |  |
| $A$12 | KeyboardText | False | False |  |

### Style Analysis: DB Cluster Cells vs Non-Cluster Cells

- Unique styles used ONLY by DB clusters: None
- Styles shared with non-cluster cells: 


### Named Range Analysis

- _xlnm.Print_Area: Sheet1!$A$1:$D$12
## Sheet Analysis

Total cells in workbook: 7
Merged cells in workbook: 0
DB clusters: 6

### DB Cluster vs Workbook Cell Analysis

| DB Address | Type | In Merge? | Has Data? | Workbook Ref |
|------------|------|-----------|-----------|--------------|
| $A$1:$B$2 | KeyboardText | False | False |  |
| $C$1:$D$2 | KeyboardText | False | False |  |
| $A$3:$D$4 | KeyboardText | False | False |  |
| $A$6:$D$7 | KeyboardText | False | False |  |
| $A$9:$D$10 | KeyboardText | False | False |  |
| $A$12 | KeyboardText | False | False |  |

### Style Analysis: DB Cluster Cells vs Non-Cluster Cells

- Unique styles used ONLY by DB clusters: None
- Styles shared with non-cluster cells: 


### Named Range Analysis

- _xlnm.Print_Area: Sheet1!$A$1:$D$12

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
