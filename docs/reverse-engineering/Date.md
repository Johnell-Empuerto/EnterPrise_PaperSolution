# Date — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Date` |
| XML `<type>` | `Date` |
| Comment line 1 | `Date` |
| Parameter class | `DateClusterParameter` |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `DateClusterParameter` (3 hits), `Date` (415 hits broad) |
| **ConMas.iReporter.UserControls.dll** | `TxtDateForm` (designer) |
| **Cimtops.R2Cluster.dll** | `Date` referenced in type list (6 hits) |
| **iReporterExcelAddInCommon.dll** | `Date` referenced (7 hits) |

## Configuration Keys

| Key | Type | Default | Values | Evidence |
|-----|------|---------|--------|----------|
| `Required` | bool | 0 | 0/1 | Common parameter |
| `DateFormat` | string | (system) | Date format string | LibConMas (188 hits) |
| `ReadOnly` | bool | 0 | 0/1 | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | UserControls |
| `DefaultValue` | string | (empty) | Date string | Common parameter |

## Designer Evidence

`Date` is a simpler variant of `Calendar`/`CalendarDate` that may show a text field with date validation rather than a full calendar picker.

## Differences from Calendar

| Aspect | Calendar / CalendarDate | Date |
|--------|------------------------|------|
| UI | Full calendar picker | Text field with date validation? |
| Parameters | DateFormat + TimeFormat | DateFormat only |
| Designer forms | `CalendarDateForm`, `CalendarDateTimeForm` | `TxtDateForm` |

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Memory:** `DateClusterParameter` class (shared with Calendar)

## Confidence Summary

| Item | Confidence |
|------|------------|
| Date as separate type exists | ★★★★☆ |
| DateFormat parameter | ★★★★★ |
| Text-based date input (vs calendar picker) | ★★★☆☆ |
