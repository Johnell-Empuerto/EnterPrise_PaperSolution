# Drawing / DrawingImage / FreeDraw — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Drawing` (or `DrawingImage` / `FreeDraw`) |
| XML `<type>` | `Drawing` / `DrawingImage` / `FreeDraw` |
| Parameter class | `DrawingImageClusterParameter` / `FreeDrawClusterParameter` / `DrawClusterParameter` |
| DLL class count | `DrawingImageClusterParameter` (1 hit), `FreeDrawClusterParameter` (1 hit), `DrawClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `DrawingImageClusterParameter`, `FreeDrawClusterParameter`, `DrawClusterParameter` |
| **LibConMas.dll (broad)** | `Drawing` (63 hits), `FreeDraw` (62 hits), `DrawingImage` (19 hits) |
| **ConMas.iReporter.UserControls.dll** | `DrawingImage`, `DrawingPinNo`, `DrawSettings`, `DrawFinished` |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether drawing is required | Common parameter |
| `ReadOnly` | bool | 0 | 0/1 | Whether drawing is read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys (from naming patterns)

| Key | Type | Description | Evidence |
|-----|------|-------------|----------|
| `DrawingMode` | enum | Drawing mode (Ink/Select/Erase) | InkCanvasEditingMode enum |
| `DrawingColor` | string | Pen color (R,G,B) | DrawingAttributes.Color |
| `DrawingWidth` | int | Pen width (px) | DrawingAttributes.Width |
| `DrawingStyle` | enum | Pen style (Solid/Dotted/etc.) | DrawingAttributes.StylusTip |
| `DrawingEraserSize` | int | Eraser size (px) | Eraser mode size |

## Runtime Controls (UserControls.dll)

- `DrawingImage` — Drawing image control
- `DrawingPinNo` — Pin number drawing control
- `DrawSettings` — Drawing settings dialog
- `DrawSettingsWidthTitle` — Width setting
- `DrawSettingsColorTitle` — Color setting
- `DrawSettingsBg` — Background setting
- `DrawSettingsOK` — OK button
- `DrawSettingsCancel` — Cancel button
- `DrawFinished` — Drawing complete event

## Relationship Between Types

| Type | Likely Purpose |
|------|----------------|
| `Drawing` / `DrawingImage` | General drawing (annotation on image) |
| `FreeDraw` | Free-form drawing (no image background) |
| `DrawingPinNo` | Pin number drawing (on image/pin map) |

All three likely share similar parameter configurations.

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Likely stored as base64 image in `rep_cluster.image_file` (similar to FreeText and Image)

## Confidence Summary

| Item | Confidence |
|------|------------|
| Drawing/FreeDraw field types exist | ★★★★★ |
| DrawingImageClusterParameter class | ★★★★★ |
| Drawing settings dialogs | ★★★★☆ |
| PinNo drawing variant | ★★★★☆ |
| Specific parameter keys | ★★☆☆☆ |
