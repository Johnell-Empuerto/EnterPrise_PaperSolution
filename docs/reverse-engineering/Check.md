# Check — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Check` |
| XML `<type>` | `Check` |
| Comment line 1 | `Check` |
| Parameter class | `CheckClusterParameter` |
| DLL class count | `CheckClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `CheckClusterParameter` class, `Check` (375 hits broad) |
| **ConMas.iReporter.UserControls.dll** | Check-related WPF controls |
| **ConMasClient.exe** | `CheckMinCluster`, `CheckMaxCluster`, `CheckAllowMinCluster`, `CheckAllowMaxCluster` |
| **Cimtops.R2Cluster.dll** | `Check` referenced (5 hits), `CheckCluster`, `CheckCaption`, `CheckMinMax`, `CheckSplitter` |
| **iReporterExcelAddInCommon.dll** | `CheckBox` (20 hits), `CheckChanged`, `CheckPredicate`, `Checked` |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Evidence |
|-----|------|---------|--------|----------|
| `Required` | bool | 0 | 0/1 | Common parameter |
| `Caption` | string | (empty) | Label text | Cimtops (12 hits) |
| `ReadOnly` | bool | 0 | 0/1 | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | UserControls |

### Hypothetical Keys

| Key | Type | Description | Evidence |
|-----|------|-------------|----------|
| `CheckStyle` | enum | Checkbox style | Naming pattern only |
| `CheckSize` | int | Checkbox size | Naming pattern only |
| `CheckColor` | string | Checkbox color | Naming pattern only |
| `DefaultValue` | bool | Default checked state | Common parameter pattern |

## Designer Validation Settings (ConMasClient.exe)

Found in ConMasClient.exe:
- `CheckMinCluster` — Minimum check count
- `CheckMaxCluster` — Maximum check count
- `CheckAllowMinCluster` — Allow minimum override
- `CheckAllowMaxCluster` — Allow maximum override

These are designer validation settings.

## Runtime Controls

The legacy runtime uses a WPF `CheckBox` control.

### Excel Add-in References
- `CheckBox` (20 hits) — CheckBox control
- `CheckChanged` (3 hits) — Check state changed event
- `CheckPredicate` (3 hits) — Check state predicate
- `Checked` (2 hits) — Checked state
- `CheckState` (2 hits) — Check state enumeration
- `CheckSplitter` — Check splitter setting

## Storage

- **Database:** `def_cluster.input_parameter` column
- **XML:** `<inputParameters>...</inputParameters>` within `<cluster>` element
- **Memory:** `CheckClusterParameter` class

## Confidence Summary

| Item | Confidence |
|------|------------|
| Check as field type exists | ★★★★★ |
| CheckClusterParameter class | ★★★★★ |
| Caption parameter | ★★★☆☆ |
| CheckBox as runtime control | ★★★★☆ |
| Validation min/max settings | ★★★☆☆ |
