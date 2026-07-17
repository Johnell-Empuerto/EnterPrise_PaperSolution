# Action — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Action` |
| XML `<type>` | `Action` |
| Parameter class | `ActionClusterParameter` |
| DLL class count | `ActionClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `ActionClusterParameter`, `Action` (316 hits broad), `ActionType` (1 hit), `ActionValue` (6 hits) |
| **ConMas.iReporter.UserControls.dll** | Action button WPF controls |
| **ConMasClient.exe** | `LblActionButton`, `TextModeActionButton`, `LblTextModeActionButton` |
| **LocalizableStrings.xml** | `Cluster.Action.DisplayValue.CommandDone`, `Cluster.Action.DisplayValue.OutputDone`, `Cluster.Action.DisplayValue.SetValueDone` |
| **Cimtops.R2Cluster.dll** | `ActionKind`, `ActionTypes` referenced |
| **iReporterExcelAddInCommon.dll** | `Action` (7 hits) |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether action is required | Common parameter |
| `ActionType` | enum | (default) | Command/Output/SetValue | Type of action | LibConMas (1 hit `ActionType`) |
| `ActionValue` | string | (empty) | Action target/value | Target for the action | LibConMas (6 hits) |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description | Evidence |
|-----|------|-------------|----------|
| `ActionLabel` | string | Button display label | Naming pattern |
| `ActionConfirm` | bool | Show confirmation dialog | LocalizableStrings confirm patterns |
| `ActionConfirmText` | string | Confirmation message text | LocalizableStrings confirm patterns |
| `ActionTarget` | string | Action target reference | Naming pattern |

## Action Types (from LocalizableStrings.xml)

The `Cluster.Action.DisplayValue.*` entries reveal three action result types:
| ActionTypes string | UI Display (EN) | Description |
|-------------------|-----------------|-------------|
| `CommandDone` | Command executed | Sends a command to server |
| `OutputDone` | Output completed | Outputs data to file/device |
| `SetValueDone` | Value set | Sets a value to another field |

## Designer Controls (ConMasClient.exe)

- `LblActionButton` — Action button label control
- `TextModeActionButton` — Text-mode action button
- `LblTextModeActionButton` — Text mode label

## Runtime Behaviors

Actions are displayed as buttons. When tapped:
1. **Command**: Sends a predefined command to the server
2. **Output**: Outputs form data to a file (PDF, CSV) or external device
3. **SetValue**: Sets a value to another cluster field

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Memory:** `ActionClusterParameter` class

## Confidence Summary

| Item | Confidence |
|------|------------|
| Action as field type | ★★★★★ |
| ActionClusterParameter class | ★★★★★ |
| ActionType / ActionValue params | ★★★☆☆ |
| Three action types (Command/Output/SetValue) | ★★★★☆ |
| Button UI in runtime | ★★★★★ |
