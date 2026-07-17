# Toggle — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Toggle` |
| XML `<type>` | `Toggle` |
| Parameter class | (Unknown — likely shared with CheckClusterParameter) |
| DLL count | `Toggle` (70 hits in LibConMas.dll, 2 hits in Cimtops.R2Cluster.dll) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `Toggle` referenced (70 hits) — heavily used |
| **Cimtops.R2Cluster.dll** | `Toggle` referenced (2 hits) |
| **iReporterExcelAddInCommon.dll** | `Toggle` referenced (2 hits) |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Evidence |
|-----|------|---------|--------|----------|
| `Required` | bool | 0 | 0/1 | Common parameter |
| `Caption` | string | (empty) | Label text | Common parameter |
| `ReadOnly` | bool | 0 | 0/1 | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | UserControls |

### Hypothetical Keys

| Key | Type | Description |
|-----|------|-------------|
| `ToggleStyle` | enum | Toggle switch style (On/Off, Yes/No) |
| `ToggleSize` | int | Toggle control size |
| `ToggleColor` | string | Toggle color |
| `DefaultValue` | bool | Default toggle state |

## Differences from Check

| Aspect | Check | Toggle |
|--------|-------|--------|
| cluster_type | `Check` | `Toggle` |
| UI | Checkbox | Toggle switch |
| Designer representation | Different property page? | Different property page? |

## Storage

- **Database:** `def_cluster.input_parameter` column
- **XML:** `<inputParameters>...</inputParameters>` within `<cluster>` element

## Confidence Summary

| Item | Confidence |
|------|------------|
| Toggle as distinct type | ★★★★☆ |
| Toggle referenced in code | ★★★★☆ |
| Parameters (specific to Toggle) | ★☆☆☆☆ — Not directly verified |
