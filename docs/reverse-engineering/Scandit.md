# Scandit — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Scandit` |
| XML `<type>` | `Scandit` |
| Parameter class | `ScanditClusterParameter` |
| DLL class count | `ScanditClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `ScanditClusterParameter`, `Scandit` (407 hits — highest count), `ScanditMode` (4 hits) |
| **ConMas.iReporter.UserControls.dll** | Scandit scanning controls |
| **Cimtops.R2Cluster.dll** | `Scandit` referenced (1 hit) |
| **iReporterExcelAddInCommon.dll** | `Scandit` referenced (1 hit) |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether scan is required | Common parameter |
| `ScanditMode` | enum | (default) | Single/Continuous | Scanner mode | LibConMas (4 hits) |
| `Language` | string | (default) | Language code | Recognition language | LibConMas (5 hits) |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description |
|-----|------|-------------|
| `ScanditFormat` | enum | Expected barcode/OCR format |
| `ScanditContinuous` | bool | Continuous scanning mode |
| `ScanditAutoScan` | bool | Auto-detect and scan |
| `ScanditBeepOnScan` | bool | Beep on successful scan |
| `ScanditShowResult` | bool | Show scanned result |
| `ScanditResultTarget` | string | Result output destination |

## Comparison with EdgeOCR and Barcode

| Aspect | Scandit | EdgeOCR | Barcode/QR |
|--------|---------|---------|------------|
| Usage | 407 hits (highest) | 395 hits | 44 / 129 hits |
| Focus | General-purpose scanner | Document edge OCR | Barcode/QR specific |
| Library | Scandit SDK (commercial) | Custom OCR engine | ZXing (open source) |

Scandit is the most-used field type in LibConMas.dll, indicating it's the primary scanning/mobile capture solution.

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Scan result stored as text in `rep_cluster.input_value`

## Confidence Summary

| Item | Confidence |
|------|------------|
| Scandit as field type | ★★★★★ |
| ScanditClusterParameter class | ★★★★★ |
| Most-used field type in LibConMas | ★★★★★ |
| ScanditMode parameter | ★★★★☆ |
| Hypothetical scan UX params | ★☆☆☆☆ |
