# PinNo / Pin — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `PinNo` |
| XML `<type>` | `PinNo` |
| Parameter class | `DrawingPinNoClusterParameter` / `PinItemTableNoClusterParameter` |
| DLL class count | `DrawingPinNoClusterParameter` (1 hit), `PinItemTableNoClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `DrawingPinNoClusterParameter`, `PinItemTableNoClusterParameter`, `PinNo` (38 hits), `PinImage` (6 hits), `PinColor` (4 hits) |
| **ConMas.iReporter.UserControls.dll** | `DrawingPinNo`, `MovePinWindow`, `PinItemTableNo` |
| **ConMasClient.exe** | `PinCooperationSelectCluster`, `PinImageCluster`, `PinItemCluster`, `LastPinTargetPage` |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether pin is required | Common parameter |
| `PinImage` | string | (default) | Image reference | Pin icon/image | LibConMas (6 hits) |
| `PinColor` | string | (default) | R,G,B | Pin color | LibConMas (4 hits) |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description |
|-----|------|-------------|
| `PinSize` | int | Pin icon size (px) |
| `PinLabel` | string | Pin display label |
| `PinLabelPosition` | enum | Label position (Top/Bottom/Left/Right) |
| `PinShape` | enum | Pin shape (Circle/Square/Pushpin) |
| `PinNo` | int | Pin number identifier |

## Designer Features (ConMasClient.exe)

- `PinCooperationSelectCluster` — Pin cooperation selection
- `PinImageCluster` — Pin image configuration
- `PinItemCluster` — Pin item configuration
- `LastPinTargetPage` — Last pin target page

## Runtime Controls (UserControls.dll)

- `DrawingPinNo` — Pin number drawing control
- `MovePinWindow` — Window for moving pins on canvas
- `PinItemTableNo` — Pin item table number

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Pin position/coordinates stored in `rep_cluster`

## Confidence Summary

| Item | Confidence |
|------|------------|
| PinNo as field type | ★★★★★ |
| DrawingPinNoClusterParameter | ★★★★★ |
| PinImage, PinColor params | ★★★★☆ |
| Pin cooperation/move UI | ★★★★☆ |
| Hypothetical size/label params | ★☆☆☆☆ |
