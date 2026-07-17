# QRCode — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `QRCode` |
| XML `<type>` | `QRCode` |
| Parameter class | `QRCodeClusterParameter` |
| DLL class count | `QRCodeClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `QRCodeClusterParameter` class, `QRCode` (129 hits broad) |
| **ConMas.iReporter.UserControls.dll** | QR scanning controls |
| **ConMasClient.exe** | `SetQrDivideToCluster` — QR code division setting |
| **Cimtops.R2Cluster.dll** | `QRCode` referenced |
| **zxing.dll** | ZXing library handles QR decoding |

## Configuration Keys

| Key | Type | Default | Values | Evidence |
|-----|------|---------|--------|----------|
| `Required` | bool | 0 | 0/1 | Common parameter |
| `ScanMode` | enum | (default) | Single/Continuous | LibConMas (9 hits) |
| `Timeout` | int | 0 | ms | LibConMas (10 hits) |
| `QRCodeFormat` | enum | (default) | QR format | Naming pattern |
| `ReadOnly` | bool | 0 | 0/1 | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | UserControls |

QRCode shares most parameters with Barcode (same scanning infrastructure via zxing.dll).

## Designer Features (ConMasClient.exe)

- `SetQrDivideToCluster` — QR code data division into cluster values
- `GetBarcodeDecompositionsUsesCluster` — Barcode/QR decomposition

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Scanned QR value stored as text in `rep_cluster.input_value`

## Confidence Summary

| Item | Confidence |
|------|------------|
| QRCode as field type | ★★★★★ |
| QRCodeClusterParameter class | ★★★★★ |
| Shared scanning infra with Barcode | ★★★★☆ |
| ZXing library | ★★★★★ |
