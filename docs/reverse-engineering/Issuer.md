# Issuer — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Issuer` |
| XML `<type>` | `Issuer` |
| Parameter class | (Unknown — system-generated field) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `Issuer` (2 hits strict) |
| **iReporterExcelAddInCommon.dll** | `Issuer` referenced (1 hit) |

## Configuration Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `ReadOnly` | bool | 1 | 0/1 | System-generated — always readonly | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

Issuer is a **system-generated field** — automatically populated with the document issuer's user name.

## Behavior

At report creation time, the system automatically fills this field with the user name of the person who created/issued the report.

## Related System-Generated Fields

| Field Type | Auto-Populated Value |
|------------|---------------------|
| `Issuer` | Document issuer's user name |
| `LoginUser` | Current login user name/ID |
| `RegistrationDate` | Report creation date/time |
| `LatestUpdateDate` | Last update date/time |
| `ExpiryDate` | Expiry date (from form settings) |

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Auto-populated with issuer user name at report creation

## Confidence Summary

| Item | Confidence |
|------|------------|
| Issuer as field type | ★★★★☆ |
| System-generated behavior | ★★★★☆ |
| Specific parameter keys | ★★☆☆☆ (none beyond common) |
