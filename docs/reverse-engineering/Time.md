# Time — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Time` |
| XML `<type>` | `Time` |
| Parameter class | `TimeClusterParameter` |
| DLL class count | `TimeClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `TimeClusterParameter` class, `TimeCalculateClusterParameter` (related) |
| **ConMas.iReporter.UserControls.dll** | `TimeTextBox`, `TimeCalculateTextBox` |
| **ConMasClient.exe** | `TxtTimeForm` (text time form), `TimeCalForm`, `TimeCalculateForm` |
| **Cimtops.R2Cluster.dll** | `Time` referenced (2 hits), `TimeCalculate` (1 hit) |
| **iReporterExcelAddInCommon.dll** | `Time` referenced (5 hits) |

## Configuration Keys

| Key | Type | Default | Values | Evidence |
|-----|------|---------|--------|----------|
| `Required` | bool | 0 | 0/1 | Common parameter |
| `TimeFormat` | string | (system) | HH:mm / HH:mm:ss | LibConMas (64 hits) |
| `ReadOnly` | bool | 0 | 0/1 | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | UserControls |
| `DefaultValue` | string | (empty) | Time string | Common parameter |

## Runtime Controls (UserControls.dll)

- `TimeTextBox` — Time input text box
- `TimeCalculateTextBox` — Time calculation text box
- `TimeLabel` — Time display label
- `TimeSlider` — Time slider control
- `SimpleTimeSelectDialog` — Time selection popup dialog

## Related Type: TimeCalculate

| Aspect | Time | TimeCalculate |
|--------|------|---------------|
| cluster_type | `Time` | `TimeCalculate` |
| Behavior | Fixed time input | Dynamic time calculation |
| Parameter class | `TimeClusterParameter` | `TimeCalculateClusterParameter` |
| Controls | `TimeTextBox` | `TimeCalculateTextBox`, `TimeCalculateForm` |

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Memory:** `TimeClusterParameter` / `TimeCalculateClusterParameter` class

## Confidence Summary

| Item | Confidence |
|------|------------|
| TimeFormat parameter | ★★★★☆ |
| TimeTextBox runtime control | ★★★★☆ |
| Time as separate type | ★★★★☆ |
| TimeCalculate as separate variant | ★★★☆☆ |
