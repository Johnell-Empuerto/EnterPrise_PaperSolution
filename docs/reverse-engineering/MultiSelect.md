# MultiSelect — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `MultiSelect` (or `MultipleChoice`) |
| XML `<type>` | `MultiSelect` / `MultipleChoice` |
| Parameter class | `MultiSelectClusterParameter` (or `MultipleChoiceNumberParameter`) |
| DLL class count | `MultiSelectClusterParameter` (1 hit), `MultipleChoiceNumberParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `MultiSelectClusterParameter`, `MultiSelect` (62 hits), `MultipleChoice` (38 hits) |
| **ConMas.iReporter.UserControls.dll** | `SimpleMultiSelectDialog`, `MultiSelectPad`, `MyListBox` |
| **ConMasClient.exe** | `MultiSelectPad`, `MultiSelectUseKeyboardCheck`, `MultiSelectEditLabelCheck` |
| **LibConMas.dll (property names)** | `MultiSelectFontValue`, `MultiSelectFontSizeValue`, `MultiSelectWeightValue`, `MultiSelectColorValue`, `MultiSelectAlignValue`, `MultiSelectVerticalAlignValue`, `MultiSelectPaddingValue`, `MultiSelectMaxValue`, `MultiSelectLines`, `MultiSelectDisplayValue`, `MultiSelectValue`, `MultiSelectInputRestrictionValue`, `MultiSelectProhibitedCharValue`, `MultiSelectExternal`, `MultiSelectReadOnly` |

## Configuration Keys

### Verified Keys (from LibConMas.dll property names)

| Key | Type | Default | Description | Evidence |
|-----|------|---------|-------------|----------|
| `Required` | bool | 0 | Whether selection is mandatory | Common parameter |
| `MultiSelectExternal` | bool | 0 | External data source | LibConMas property name (3 hits) |
| `MultiSelectReadOnly` | bool | 0 | Read-only mode | LibConMas property name (3 hits) |
| `MultiSelectValue` | string | (empty) | Current/selected value | LibConMas property name (3 hits) |
| `MultiSelectDisplayValue` | string | (empty) | Display text | LibConMas property name (3 hits) |
| `MultiSelectFontValue` | string | Arial | Font family | LibConMas property name (3 hits) |
| `MultiSelectFontSizeValue` | int | 11 | Font size | LibConMas property name (3 hits) |
| `MultiSelectWeightValue` | enum | Normal | Font weight | LibConMas property name (3 hits) |
| `MultiSelectColorValue` | string | 0,0,0 | Text color (R,G,B) | LibConMas property name (3 hits) |
| `MultiSelectAlignValue` | enum | Center | Text alignment | LibConMas property name (3 hits) |
| `MultiSelectVerticalAlignValue` | int | 2 | Vertical alignment | LibConMas property name (3 hits) |
| `MultiSelectPaddingValue` | int | 0 | List item padding | LibConMas property name (3 hits) |
| `MultiSelectMaxValue` | int | 0 | Max selections allowed | LibConMas property name (3 hits) |
| `MultiSelectLines` | int | 0 | Lines to display | LibConMas property name (3 hits) |
| `MultiSelectUseKeyboardCheck` | bool | 0 | Enable keyboard input | LibConMas property name (3 hits) |
| `MultiSelectEditLabelCheck` | bool | 0 | Allow label editing | LibConMas property name (3 hits) |
| `MultiSelectInputRestrictionValue` | enum | None | Input restriction | LibConMas property name (3 hits) |
| `MultiSelectProhibitedCharValue` | string | (empty) | Prohibited characters | LibConMas property name (3 hits) |

### Common Parameters (shared with all field types)

| Key | Type | Default | Evidence |
|-----|------|---------|----------|
| `ReadOnly` | bool | 0 | ClusterParameter base |
| `Hidden` | bool | 0 | UserControls |
| `DefaultValue` | string | (empty) | Common parameter |

## Runtime Controls (UserControls.dll)

- `SimpleMultiSelectDialog` — Multi-select popup dialog
- `MultiSelectPad` — Multi-select input pad
- `MyListBox` — Custom list box control

## Designer Property Pages (ConMasClient.exe)

- `MultiSelectPad` — Settings pad
- `MultiSelectUseKeyboardCheck` — Keyboard input checkbox
- `MultiSelectEditLabelCheck` — Label edit checkbox

## Evidence: Property Name Pattern

The property names with `Value` suffix (e.g., `MultiSelectFontValue`) suggest runtime values that override the static `Font` key. This is a pattern found only in MultiSelect and ChoiceNumber field types.

## Storage

- **Database:** `def_cluster.input_parameter` column
- **XML:** `<inputParameters>...</inputParameters>` within `<cluster>` element
- **Memory:** `MultiSelectClusterParameter` class

## Confidence Summary

| Item | Confidence |
|------|------------|
| MultiSelect cluster type exists | ★★★★★ |
| 18 configuration properties (with Value suffix) | ★★★★☆ |
| Runtime dialogs exist | ★★★★☆ |
| Exact mapping of Value-suffixed keys | ★★★☆☆ |
