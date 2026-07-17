# EdgeOCR — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `EdgeOCR` |
| XML `<type>` | `EdgeOCR` |
| Parameter class | `EdgeOCRClusterParameter` |
| DLL class count | `EdgeOCRClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `EdgeOCRClusterParameter` (1 hit), `EdgeOCR` (395 hits broad), `EdgeOCRMethods` (9 hits), `EdgeOCRCluster` (12 hits) |
| **ConMas.iReporter.UserControls.dll** | EdgeOCR scanning controls |
| **ConMasClient.exe** | `EdgeOCRScanSettingControl`, `EdgeOCRMappingCluster`, `EdgeOCRMappingSettingOutputCluster`, `EdgeOCROutputCluster`, `DividedEdgeOCRWindow`, `IsEdgeOCRCluster`, `RemoveInvalidEdgeOCRCluster`, `GetEdgeOCRCluster`, `GetEdgeOCRClusterUsesCluster`, `SetDpmDivideToCluster` |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether OCR is required | Common parameter |
| `Language` | string | (default) | Language code | OCR language | LibConMas (5 hits) |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description |
|-----|------|-------------|
| `EdgeOCRMode` | enum | OCR mode (Auto/Manual) |
| `EdgeOCRFormat` | enum | Output format (Text/Number/Date) |
| `EdgeOCRLanguage` | string | OCR recognition language |
| `EdgeOCRRegion` | string | Region of interest (ROI) |
| `EdgeOCRContinuous` | bool | Continuous scanning |
| `EdgeOCRAutoScan` | bool | Auto-detect and scan |
| `EdgeOCRBeepOnScan` | bool | Beep on successful scan |
| `EdgeOCRShowResult` | bool | Show OCR result |
| `EdgeOCRResultTarget` | string | Result output destination |

## Designer Features (ConMasClient.exe)

### Scan Settings
- `EdgeOCRScanSettingControl` — OCR scan settings control
- `SetDpmDivideToCluster` — DPM (dots per millimeter) division to cluster

### Mapping/Output
- `EdgeOCRMappingCluster` — OCR field mapping
- `EdgeOCRMappingSettingOutputCluster` — Mapping output settings
- `EdgeOCROutputCluster` — OCR output cluster
- `EdgeOCRMappingSettingWindow` — Mapping settings window

### Validation
- `IsEdgeOCRCluster` — Check if cluster is EdgeOCR type
- `RemoveInvalidEdgeOCRCluster` — Remove invalid OCR clusters
- `GetEdgeOCRCluster` — Get OCR cluster configuration
- `GetEdgeOCRClusterUsesCluster` — OCR cluster dependencies

### Divided Scanning
- `DividedEdgeOCRWindow` — Divided OCR scanning window
- `SetDpmDivideToCluster` — DPM division setting

## Comparison with Scandit

| Aspect | EdgeOCR | Scandit |
|--------|---------|---------|
| cluster_type | `EdgeOCR` | `Scandit` |
| Purpose | OCR (optical character recognition) on document edges | General-purpose barcode/QR scanning |
| Parameter class | `EdgeOCRClusterParameter` | `ScanditClusterParameter` |
| Designer features | Extensive mapping, divided scanning, DPM | (Less designer integration) |
| DLL hits | 395 (LibConMas) | 407 (LibConMas) |

Both are heavily used in LibConMas.dll with similar usage frequency.

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** OCR result stored as text in `rep_cluster.input_value`

## Confidence Summary

| Item | Confidence |
|------|------------|
| EdgeOCR as field type | ★★★★★ |
| EdgeOCRClusterParameter class | ★★★★★ |
| Extensive designer controls | ★★★★★ |
| OCR language parameter | ★★★★☆ |
| Mapping/output infrastructure | ★★★★☆ |
| Hypothetical EdgeOCR-specific params | ★★☆☆☆ |
