# GPS / Location — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `GPS` |
| XML `<type>` | `GPS` |
| Parameter class | `GpsClusterParameter` |
| DLL class count | `GpsClusterParameter` (1 hit in LibConMas.dll) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `GpsClusterParameter` class, `GPS` (6 hits strict) |
| **ConMas.iReporter.UserControls.dll** | `Location` (5 hits), `Coordinates` (1 hit) |
| **LocalizableStrings.xml** | `Settings.Detail.Cluster.GpsDistanceFilter` with values: 10m, 100m, 1000m |
| **iReporterExcelAddInCommon.dll** | `GPS` (4 hits), `Location` (3 hits) |
| **Cimtops.R2Cluster.dll** | `GPS` (1 hit) |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether GPS is required | Common parameter |
| `GPSDistanceFilter` | int | (default) | 10/100/1000 (meters) | GPS update distance filter | LocalizableStrings.xml |
| `Timeout` | int | 0 | ms | GPS acquisition timeout | LibConMas (10 hits) |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description |
|-----|------|-------------|
| `GPSUpdateInterval` | int | Location update interval (ms) |
| `GPSAccuracy` | int | Required accuracy (meters) |
| `GPSAutoStop` | bool | Auto-stop GPS after acquisition |
| `GPSShowMap` | bool | Show map UI |
| `GPSShowCoordinates` | bool | Show lat/lng coordinates |

## Settings UI (from LocalizableStrings.xml)

`Settings.Detail.Cluster.GpsDistanceFilter` with options:
- 10m (10 meters)
- 100m (100 meters)
- 1000m (1000 meters)

This filter determines when GPS updates are sent — only when the device moves beyond the specified distance.

## Data Stored

GPS fields likely store:
- Latitude
- Longitude
- Accuracy
- Timestamp
- (Possibly address)

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** GPS coordinates stored as text in `rep_cluster.input_value`

## Confidence Summary

| Item | Confidence |
|------|------------|
| GPS as field type | ★★★★★ |
| GpsClusterParameter class | ★★★★★ |
| Distance filter (10/100/1000m) | ★★★★★ |
| Location data stored | ★★★☆☆ |
| Hypothetical GPS parameter keys | ★☆☆☆☆ |
