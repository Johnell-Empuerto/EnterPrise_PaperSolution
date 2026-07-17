# LoginUser — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `LoginUser` |
| XML `<type>` | `LoginUser` |
| Parameter class | `LoginUserClusterParameter` |
| DLL class count | `LoginUserClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `LoginUserClusterParameter`, `LoginUser` (69 hits), `LoginUserId` (7 hits) |
| **ConMasClient.exe** | `LoginUserDialog` |
| **Cimtops.R2Cluster.dll** | `LoginUser` referenced (1 hit) |
| **iReporterExcelAddInCommon.dll** | `LoginUser` (not directly found but related) |

## Configuration Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `ReadOnly` | bool | 1 | 0/1 | System-generated — always readonly | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

LoginUser is a **system-generated field** — it is automatically populated with the current user's information and has minimal configuration.

## Behavior

LoginUser automatically captures:
1. **LoginUserId** — The current user's login ID
2. **LoginUserName** / **LoginUser** — The current user's display name

This field type is read-only at runtime since values are system-generated.

## Designer UI (ConMasClient.exe)

`LoginUserDialog` — Dialog for configuring LoginUser field properties (likely minimal since auto-populated).

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Auto-populated with current user info at report creation time

## Confidence Summary

| Item | Confidence |
|------|------------|
| LoginUser as field type | ★★★★★ |
| LoginUserClusterParameter class | ★★★★★ |
| System-generated (auto-populated) | ★★★★☆ |
| LoginUserId / LoginUserName storage | ★★★★☆ |
| Read-only at runtime | ★★★★☆ |
