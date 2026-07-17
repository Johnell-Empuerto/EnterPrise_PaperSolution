# Approve — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Approve` |
| XML `<type>` | `Approve` |
| Comment line 1 | `Approve` |
| Parameter class | `ApproveClusterParameter` |
| DLL class count | `ApproveClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `ApproveClusterParameter`, `Approve` (41 hits), `ApproveSign` (5 hits), `ApproveRequired` (6 hits) |
| **ConMas.iReporter.UserControls.dll** | `SimpleApproveDialog`, `ApproveButtonTitle`, `ApproveCreateCommentTitle`, `ApproveCreateDateTitle`, `ApproveCreateDateValue`, `ApproveCreateCommentValue`, `ApproveUserIdValue`, `ApproveUserNameValue` |
| **DefinitionDetailBiz.xml** | `AND cluster_type = 'Approve'` — filtered query |
| **LocalizableStrings.xml** | `Approve.Type.Editting`, `Approve.Type.Unedited` |
| **ConMasClient.exe** | Accept/Reject approve buttons |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether approval is required | Common parameter |
| `ApproveRequired` | bool | 1 | 0/1 | Whether signature is required | LibConMas (6 hits) |
| `ApproveSign` | bool | 1 | 0/1 | Show signature field | LibConMas (5 hits) |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Verified from UserControls.dll

| Key | Type | Description | Evidence |
|-----|------|-------------|----------|
| `ApproveDateVisible` | bool | Show approval date | `ApproveCreateDateValue`, `ApproveCreateDateTitle` |
| `ApproveCommentVisible` | bool | Show approval comment | `ApproveCreateCommentValue`, `ApproveCreateCommentTitle` |

### Hypothetical Keys

| Key | Type | Description |
|-----|------|-------------|
| `ApproveType` | enum | Approval type (Sign/Comment/Full) |
| `ApproveLabel` | string | Approval button label |

## Designer Evidence (DefinitionDetailBiz.xml)

```xml
<approve>
  SELECT cluster_name AS "name" FROM def_cluster
  WHERE def_sheet_id IN (...)
    AND cluster_type = 'Approve'
  ORDER BY def_sheet_id, cluster_id
</approve>
```

This confirms `Approve` is a distinct cluster_type used in approval flow queries.

## Runtime Data Stored

From UserControls.dll, the approval stores:
- `ApproveCreateDateValue` — Date/time of approval
- `ApproveCreateCommentValue` — Comment/remark text
- `ApproveUserIdValue` — Approver's user ID
- `ApproveUserNameValue` — Approver's user name

## States (from LocalizableStrings.xml)
- `Approve.Type.Editting` — "Editing" (not yet approved)
- `Approve.Type.Unedited` — "Unedited" (no action taken)

## Runtime Dialog
- `SimpleApproveDialog` — Approval confirmation dialog
- `NoApproveButton` — Reject/No approval button
- `ApproveButtonTitle` — Approve button text

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Approval data (user, date, comment, signature) stored across `rep_cluster` fields

## Confidence Summary

| Item | Confidence |
|------|------------|
| Approve as field type | ★★★★★ |
| ApproveClusterParameter class | ★★★★★ |
| ApproveRequired, ApproveSign params | ★★★★☆ |
| Date/comment visibility params | ★★★★☆ |
| Runtime dialog (UserControls) | ★★★★★ |
| Approval data stored (user/date/comment) | ★★★★☆ |
