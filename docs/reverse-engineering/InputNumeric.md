# InputNumeric — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `InputNumeric` |
| XML `<type>` | `InputNumeric` |
| Comment line 1 | `InputNumeric` |
| Parameter class | `InputNumericClusterParameter` / `NumericClusterParameter` |
| DLL class count | `InputNumericClusterParameter` (1 hit), `NumericClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `InputNumericClusterParameter`, `NumericClusterParameter` classes |
| **ConMas.iReporter.UserControls.dll** | Numeric input WPF controls, number pad popup |
| **DB Template 548** | Cluster 1 (`Machine_Output`) is InputNumeric type |
| **LocalizableStrings.xml** | `InputNumeric.Error.Title`, `InputNumeric.Error.Minimum`, `InputNumeric.Error.Maximum` |
| **ConMasClient.exe** | `InputNumericMinCluster`, `InputNumericMaxCluster`, `InputNumericAllowMinCluster`, `InputNumericAllowMaxCluster` |
| **Cimtops.R2Cluster.dll** | `Numeric` referenced in type list |

## Configuration Keys

### Verified Keys (strong DLL + DB evidence)

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether input is mandatory | Common parameter pattern |
| `MinValue` | decimal | (none) | Any decimal | Minimum allowed value | LibConMas (124 hits), ConMasClient.exe (`InputNumericMinCluster`) |
| `MaxValue` | decimal | (none) | Any decimal | Maximum allowed value | LibConMas (139 hits), ConMasClient.exe (`InputNumericMaxCluster`) |
| `Format` | string | (none) | Number format | Display format | LibConMas (21 hits) |
| `DecimalPlaces` | int | 0 | 0+ | Number of decimal places | LibConMas (1 hit `DecimalPlaces`) |

### Visual Keys (from common ClusterParameter base)

| Key | Type | Default | Values | Evidence |
|-----|------|---------|--------|----------|
| `Font` | string | Arial | Font name | LibConMas (21 hits) |
| `FontSize` | int | 11 | Font size | LibConMas (57 hits) |
| `Weight` | enum | Normal | Normal/Bold | LibConMas (23 hits) |
| `Color` | string | 0,0,0 | R,G,B | LibConMas (82 hits) |
| `Align` | enum | Left/Center/Right | Alignment | LibConMas (46 hits) |
| `VerticalAlignment` | int | 0=Top/1=Center/2=Bottom | V-Align | LibConMas (9 hits) |
| `Padding` | int | 0 | Padding px | LibConMas (12 hits) |
| `ReadOnly` | bool | 0 | 0/1 | ClusterParameter base (117 hits) |
| `Hidden` | bool | 0 | 0/1 | UserControls (33 hits) |
| `DefaultValue` | string | (empty) | Default numeric value | LibConMas (9 hits) |

### Validation Keys (from ConMasClient.exe property pages)

| Key | Type | Default | Evidence |
|-----|------|---------|----------|
| `AllowMinCluster` | (validation) | — | ConMasClient.exe validation cluster names |
| `AllowMaxCluster` | (validation) | — | ConMasClient.exe validation cluster names |
| `CheckMinCluster` | (validation) | — | ConMasClient.exe validation cluster names |
| `CheckMaxCluster` | (validation) | — | ConMasClient.exe validation cluster names |

These appear to be designer-only validation settings, not input_parameter keys.

## Designer Property Pages (ConMasClient.exe)

The following property page/window names were found in ConMasClient.exe:
- `InputNumericMinCluster` — Minimum value setting
- `InputNumericMaxCluster` — Maximum value setting
- `InputNumericAllowMinCluster` — Allow minimum override
- `InputNumericAllowMaxCluster` — Allow maximum override
- `SimpleNumberInputDialog` — Number input dialog
- `NumberPopup` — Number input popup

## Runtime Behavior

The legacy runtime uses a WPF `TextBox` with numeric-only input validation or a custom number pad popup.

### Validation Messages (from LocalizableStrings.xml)
- `InputNumeric.Error.Title` — "Input Error"
- `InputNumeric.Error.Minimum` — "Please enter a value greater than {0}"
- `InputNumeric.Error.Maximum` — "Please enter a value less than {0}"

### Runtime Controls
- `NumberPopup` — Popup number pad for tablet input
- `NumberCallBack` — Callback for number input completion
- `SelectNumericCounterDialog` — Counter dialog (in UserControls.dll)

## Comment Format (Excel Add-In)

```
Line 0: Field Name (e.g., "Machine_Output")
Line 1: InputNumeric
Line 2: Cluster Index (sort order, optional)
```

## Storage

- **Database:** `def_cluster.input_parameter` column
- **XML:** `<inputParameters>...</inputParameters>` within `<cluster>` element
- **Memory:** `InputNumericClusterParameter` / `NumericClusterParameter` class

## Confidence Summary

| Item | Confidence |
|------|------------|
| MinValue / MaxValue parameters | ★★★★★ |
| Visual parameters (Font, Color, etc.) | ★★★★☆ |
| Format / DecimalPlaces | ★★★☆☆ |
| Validation settings | ★★★☆☆ |
| Number popup runtime control | ★★★★☆ |
