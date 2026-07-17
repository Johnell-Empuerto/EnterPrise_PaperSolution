# Calculation — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Calculation` |
| XML `<type>` | `Calculation` |
| Parameter class | (Unknown — possibly shared with `TimeCalculateClusterParameter`) |
| DLL class count | `Calculation` (3 hits in LibConMas strict search) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `Calculation` (3 hits strict, related `TimeCalculate` has its own parameter class) |
| **ConMas.iReporter.UserControls.dll** | Calculation-related WPF controls |
| **ConMasClient.exe** | `CalculateDateForm`, `CalculateMaxCluster`, `CalculateMinCluster`, `CalculateAllowMaxCluster`, `CalculateAllowMinCluster`, `TimeCalForm`, `TimeCalculateForm` |
| **iReporterExcelAddInCommon.dll** | `Calculation` referenced (2 hits) |
| **IReporter.Calculate.dll** | Calculation engine library (23KB) |

## Configuration Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether result is required | Common parameter |
| `ReadOnly` | bool | 1 | 0/1 | Calculated fields are read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description |
|-----|------|-------------|
| `CalculationFormula` | string | Formula expression |
| `CalculationDecimal` | int | Decimal places for result |
| `CalculationFormat` | string | Result display format |

## Related Library

`IReporter.Calculate.dll` (23KB) — Dedicated calculation engine for evaluating formulas.

## Designer Validation (ConMasClient.exe)

- `CalculateMaxCluster` — Maximum value validation
- `CalculateMinCluster` — Minimum value validation
- `CalculateAllowMaxCluster` — Allow max override
- `CalculateAllowMinCluster` — Allow min override
- `CalculateDateForm` — Calculation date form

## Related: TimeCalculate

`TimeCalculate` is a related time-specific calculation variant with its own parameter class (`TimeCalculateClusterParameter`).

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Calculated result stored as text in `rep_cluster.input_value` (auto-populated, read-only)

## Confidence Summary

| Item | Confidence |
|------|------------|
| Calculation as field type | ★★★★☆ |
| IReporter.Calculate.dll | ★★★★★ |
| Designer validation clusters | ★★★★☆ |
| Formula/format parameters | ★☆☆☆☆ |
