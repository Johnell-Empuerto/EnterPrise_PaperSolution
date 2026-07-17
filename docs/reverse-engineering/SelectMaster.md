# SelectMaster — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `SelectMaster` |
| XML `<type>` | `SelectMaster` |
| Parameter class | `MasterClusterParameter` |
| DLL class count | `MasterClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `MasterClusterParameter`, `SelectMaster` (83 hits broad), `MasterValue` (3 hits), `MasterType` (3 hits), `MasterKey` (7 hits), `MasterInfoService` |
| **ConMas.iReporter.UserControls.dll** | Master selection WPF controls |
| **ConMasClient.exe** | `SelectMaster` related property pages |
| **Cimtops.R2Cluster.dll** | `SelectMaster` referenced (1 hit) |
| **iReporterExcelAddInCommon.dll** | `SelectMaster` referenced (1 hit) |
| **CustomMaster.xml** | `master_label_entity` — master label definition table |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether selection is required | Common parameter |
| `MasterType` | enum | (default) | Custom/System/Label | Master data type | LibConMas (3 hits) |
| `MasterValue` | string | (empty) | Selected value | Current selected master value | LibConMas (3 hits) |
| `MasterKey` | string | (empty) | Key field | Master key field name | LibConMas (7 hits) |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description |
|-----|------|-------------|
| `MasterSource` | string | Master data source (table/view) |
| `MasterDisplay` | string | Display field name |
| `MasterFilter` | string | Filter expression |
| `MasterCache` | bool | Cache master data locally |
| `MasterAutoComplete` | bool | Auto-complete on input |
| `MasterSort` | string | Sort expression |
| `MasterDisplayFormat` | string | Display format pattern |

## Service Reference

`MasterInfoService` in LibConMas.dll — provides master data lookup service.

## Database Schema (CustomMaster.xml)

The `master_label_entity` table stores master label definitions:
```sql
CREATE TABLE master_label_entity (
    label_id, upper_label_id, name, display_number,
    common, group_id, icon_id, indentation_level
)
```

This table serves as a hierarchical master data source for `SelectMaster` fields.

## Runtime Behavior

SelectMaster provides a drop-down/list selection from a master data source. The user selects a value which is stored as the cluster value.

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Selected master value stored in `rep_cluster.input_value`

## Confidence Summary

| Item | Confidence |
|------|------------|
| SelectMaster as field type | ★★★★★ |
| MasterClusterParameter class | ★★★★★ |
| MasterType, MasterValue, MasterKey | ★★★★☆ |
| master_label_entity DB table | ★★★★☆ |
| MasterInfoService | ★★★☆☆ |
