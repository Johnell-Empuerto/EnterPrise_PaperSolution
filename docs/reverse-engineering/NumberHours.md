# NumberHours — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `NumberHours` |
| XML `<type>` | `NumberHours` |
| Parameter class | `NumberHoursClusterParameter` |
| DLL class count | `NumberHoursClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `NumberHoursClusterParameter`, `NumberHours` (60 hits), `NumberHoursFormat` (9 hits), `NumberHoursDecimal` (3 hits), `NumberHoursMax` (3 hits) |
| **ConMas.iReporter.UserControls.dll** | `NumberHour`, `NumberHourSplitChar` (from `InputParameterNumberHourSplitChar`), `NumberHoursTimeForm` |
| **ConMasClient.exe** | `NumberHoursForm`, `NumberHoursTimeForm`, `NumberHoursMaxCluster`, `TxtNumberHoursForm` |
| **LocalizableStrings.xml** | `NumberHours.Error.Title` |
| **Cimtops.R2Cluster.dll** | `NumberHours` referenced (2 hits) |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether input is required | Common parameter |
| `NumberHoursFormat` | string | (default) | HH:mm / decimal | Hours display format | LibConMas (9 hits) |
| `NumberHoursDecimal` | int | 0 | 0+ | Decimal places for decimal format | LibConMas (3 hits) |
| `NumberHoursMax` | decimal | (none) | Max hours | Maximum hours value | LibConMas (3 hits), ConMasClient `NumberHoursMaxCluster` |
| `NumberHoursMin` | decimal | (none) | Min hours | Minimum hours value | Pattern (max exists, min implied) |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description | Evidence |
|-----|------|-------------|----------|
| `NumberHoursAutoCalc` | bool | Auto-calculate from date/time fields | Common calculation pattern |
| `NumberHoursRoundUp` | bool | Round up to nearest | Common rounding |
| `NumberHoursRoundDown` | bool | Round down to nearest | Common rounding |
| `NumberHoursRoundMinute` | int | Round to nearest N minutes | Common rounding |

### Hidden Key
| Key | Type | Description | Evidence |
|-----|------|-------------|----------|
| `NumberHourSplitChar` | string | Character used to split hour/minute input | UserControls (`InputParameterNumberHourSplitChar`) |

## Runtime Controls (UserControls.dll)

- `NumberHour` — Number hour input control
- `NumberHourSplitChar` — Split character for hour/minute input
- `TxtNumberHoursForm` — Text form for hours input
- `NumberHoursTimeForm` — Time form for hours input

## Designer Property Pages (ConMasClient.exe)

- `NumberHoursForm` — Number hours settings form
- `NumberHoursTimeForm` — Number hours time form
- `NumberHoursMaxCluster` — Maximum hours validation
- `TxtNumberHoursForm` — Text hours form

## Validation Error (LocalizableStrings.xml)
- `NumberHours.Error.Title` — "Input Error" title

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Hours value stored as text in `rep_cluster.input_value`

## Confidence Summary

| Item | Confidence |
|------|------------|
| NumberHours as field type | ★★★★★ |
| NumberHoursClusterParameter class | ★★★★★ |
| Format, Decimal, Max/Min params | ★★★★☆ |
| NumberHourSplitChar hidden param | ★★★★☆ |
| Auto-calc and rounding params | ★☆☆☆☆ |
