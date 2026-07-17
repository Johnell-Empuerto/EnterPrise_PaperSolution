# Barcode — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Barcode` |
| XML `<type>` | `Barcode` |
| Parameter class | `CodeReaderClusterParameter` (shared with QRCode) |
| DLL class count | `CodeReaderClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `CodeReaderClusterParameter`, `Barcode` (44 hits broad) |
| **ConMas.iReporter.UserControls.dll** | Barcode scanning controls (zxing-based) |
| **zxing.dll** | ZXing barcode decoding library (442KB) |
| **LibConMas.dll (params)** | `ScanMode` (9 hits), `Timeout` (10 hits), `Language` (5 hits) |
| **ConMasClient.exe** | `GetBarcodeDecompositionsUsesCluster` — barcode decomposition |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether scan is required | Common parameter |
| `BarcodeFormat` | enum | (default) | Code128/QR/EAN13/etc. | Expected barcode format | LibConMas broad search (44 hits) |
| `ScanMode` | enum | (default) | Single/Continuous | Scanner mode | LibConMas (9 hits) |
| `Timeout` | int | 0 | ms | Scan timeout | LibConMas (10 hits) |
| `Language` | string | (default) | Language code | OCR language | LibConMas (5 hits) |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description |
|-----|------|-------------|
| `BeepOnScan` | bool | Play beep on successful scan |
| `VibrateOnScan` | bool | Vibrate on successful scan |
| `ContinuousScan` | bool | Continuous scanning mode |
| `AutoScan` | bool | Auto-detect and scan |
| `ShowResult` | bool | Show scanned result |
| `ResultTarget` | string | Result output destination |
| `BarcodeMode` | enum | Barcode scanning mode |

## Runtime Library

The runtime uses **ZXing** (`zxing.dll`) for barcode/QR code decoding:
- `zxing.dll` — Core ZXing library (442KB)
- `zxing.presentation.dll` — ZXing WPF presentation library

## Related Type: QRCode

| Aspect | Barcode | QRCode |
|--------|---------|--------|
| cluster_type | `Barcode` | `QRCode` |
| Parameter class | `CodeReaderClusterParameter` | `QRCodeClusterParameter` |
| Library | zxing.dll | zxing.dll |
| Format | 1D barcodes | 2D QR codes |

Both likely use the same scanning infrastructure but with different format defaults.

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Scanned value stored as text in `rep_cluster.input_value`

## Confidence Summary

| Item | Confidence |
|------|------------|
| Barcode as field type | ★★★★★ |
| CodeReaderClusterParameter class | ★★★★★ |
| ZXing library used | ★★★★★ |
| ScanMode, Timeout parameters | ★★★★☆ |
| BarcodeFormat parameter | ★★★☆☆ |
| Hypothetical scan UX params | ★☆☆☆☆ |
