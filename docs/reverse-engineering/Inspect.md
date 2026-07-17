# Inspect — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Inspect` |
| XML `<type>` | `Inspect` |
| Parameter class | `InspectClusterParameter` |
| DLL class count | `InspectClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `InspectClusterParameter`, `Inspect` (22 hits broad) |
| **Cimtops.R2Cluster.dll** | `Inspect` referenced (5 hits) |
| **iReporterExcelAddInCommon.dll** | `Inspect` referenced (1 hit) |

## Configuration Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether inspection is required | Common parameter |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description |
|-----|------|-------------|
| `InspectType` | enum | Inspection result type (Pass/Fail/OK/NG) |
| `InspectResult` | string | Inspection result options |
| `InspectLabel` | string | Inspection field label |

## Relationship with Approve

| Aspect | Inspect | Approve |
|--------|---------|---------|
| cluster_type | `Inspect` | `Approve` |
| Purpose | Inspection/check result | Approval/signature |
| Parameter class | `InspectClusterParameter` | `ApproveClusterParameter` |
| Designer features | Limited | Extensive (dialog, sign, date, comment) |

Inspect is likely a simpler variant of Approve, used for inspection checkpoints rather than formal approval.

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Inspection result stored as text in `rep_cluster.input_value`

## Confidence Summary

| Item | Confidence |
|------|------------|
| Inspect as field type | ★★★★★ |
| InspectClusterParameter class | ★★★★★ |
| Relationship with Approve | ★★★☆☆ |
| Specific parameter keys | ★☆☆☆☆ |
