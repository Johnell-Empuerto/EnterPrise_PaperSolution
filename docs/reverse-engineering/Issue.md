# Issue — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Issue` |
| XML `<type>` | `Issue` |
| Parameter class | (Unknown — shared with `RegistrationDate` or `Issue`-specific) |
| DLL class count | `Issue` (4 hits in LibConMas broad search) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `Issue` (4 hits broad), `Issuer` (2 hits) |
| **Cimtops.R2Cluster.dll** | `Issue` referenced (1 hit) |
| **iReporterExcelAddInCommon.dll** | `Issue` referenced (1 hit) |

## Configuration Keys

| Key | Type | Default | Description | Evidence |
|-----|------|---------|-------------|----------|
| `Required` | bool | 0 | Whether issue is required | Common parameter |
| `ReadOnly` | bool | 0 | Whether read-only (may be 1 if auto-generated) | ClusterParameter base |
| `Hidden` | bool | 0 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description |
|-----|------|-------------|
| `IssueType` | enum | Issue type |
| `IssueAutoGenerate` | bool | Auto-generate issue number |
| `IssuePrefix` | string | Issue number prefix |
| `IssueNumber` | int | Current issue number |
| `IssueDigits` | int | Number of digits for issue number |
| `IssueDate` | bool | Show issue date |

## Related Types

| Aspect | Issue | Issuer | RegistrationDate |
|--------|-------|--------|------------------|
| cluster_type | `Issue` | `Issuer` | `RegistrationDate` |
| Purpose | Issue number/tracking | Issuing person name | Registration timestamp |
| Auto-populated | Possibly (if auto-generate) | Yes (system user) | Yes (system time) |

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Issue value stored as text in `rep_cluster.input_value`

## Confidence Summary

| Item | Confidence |
|------|------------|
| Issue as field type | ★★★★☆ |
| Minimal parameter evidence | ★★☆☆☆ |
| Connection to Issuer/RegistrationDate | ★★★☆☆ |
