# Image — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Image` |
| XML `<type>` | `Image` |
| Parameter class | `ImageClusterParameter` |
| DLL class count | `ImageClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `ImageClusterParameter` class, `Image` (365 hits broad), `ImageSize` (44 hits), `ImageFormat` (7 hits), `ImageSource` (2 hits) |
| **ConMas.iReporter.UserControls.dll** | `CameraControl`, `CameraDialog`, `PhotoLibrary`, `ShowCameraControl`, `ShowPvCameraWindow`, `ImageReadyq` |
| **LocalizableStrings.xml** | `MainMenu.Title.Camera`, `MainMenu.Title.PhotoAlbum`, `Settings.ImageCorrection`, `Alert.Message.FailToCaptureImage` |
| **ClusterImageBiz.xml** | `cluster_type IN ('Image','FreeText')` — stores image data for both types |
| **Db evidence** | `rep_cluster.image_file` stores base64 image data |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether image is required | Common parameter |
| `ImageSize` | int/string | (default) | Size constraint | Max image size | LibConMas (44 hits) |
| `ImageFormat` | enum | (default) | JPG/PNG/BMP | Output image format | LibConMas (7 hits), UserControls (13 hits) |
| `ImageSource` | enum | (default) | Camera/Album/Both | Image source selection | LibConMas (2 hits), UserControls DLL names |
| `Resolution` | int | (default) | DPI value | Image resolution | UserControls (53 hits) |
| `Compression` | int | (default) | Compression level | Image compression quality | UserControls (8 hits) |
| `ReadOnly` | bool | 0 | 0/1 | Whether image is read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether field is hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description | Evidence |
|-----|------|-------------|----------|
| `ImageQuality` | int | JPEG quality % (1-100) | Pattern from ImageFormat/Compression |
| `CameraType` | enum | Front/Back camera | LocalizableStrings, menu titles |
| `AutoCapture` | bool | Auto-capture on open | Pattern from other auto params |
| `MaxCount` | int | Max number of images | Pattern from MultiSelectMaxValue |
| `MinCount` | int | Min number of images | Pattern from InputNumeric MinValue |

## Runtime Controls (UserControls.dll)

- `CameraControl` — Camera viewfinder control
- `CameraDialog` — Camera capture dialog
- `ShowCameraControl` — Shows camera control
- `ShowPvCameraWindow` — Preview camera window
- `PhotoLibrary` — Photo album/image gallery picker
- `ImageReadyq` — Image ready callback

## Designer Property Pages

- `EvidenceImageCluster` — Evidence image settings (3 hits in ConMasClient.exe)
- `GetEvidenceImageCluster` — Evidence image retrieval
- `ImageSize` — Image size setting
- `ImagePath` — Image path

## Data Storage

Images are stored in `rep_cluster.image_file` as base64-encoded binary.

From `ClusterImageBiz.xml`:
```sql
UPDATE rep_cluster SET image_file = :img_base
WHERE cluster_type IN ('Image','FreeText')
```

This confirms both `Image` and `FreeText` types can store image data.

## Menu Items (from LocalizableStrings.xml)

- `MainMenu.Title.Camera` — Camera menu
- `MainMenu.Title.PhotoAlbum` — Photo album menu
- `Settings.ImageCorrection` — Image correction setting
- `Settings.UploadQuality` — Upload quality setting

## Confidence Summary

| Item | Confidence |
|------|------------|
| Image as field type | ★★★★★ |
| ImageClusterParameter class | ★★★★★ |
| Camera / Photo UI dialogs | ★★★★★ |
| ImageSize, ImageFormat params | ★★★★☆ |
| ImageSource, Resolution, Compression | ★★★☆☆ |
| Base64 storage in rep_cluster | ★★★★★ |
