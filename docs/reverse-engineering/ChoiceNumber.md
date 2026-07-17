# ChoiceNumber — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `ChoiceNumber` |
| XML `<type>` | `ChoiceNumber` |
| Comment line 1 | `ChoiceNumber` |
| Parameter class | `MultipleChoiceNumberParameter` |
| DLL class count | `MultipleChoiceNumberParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `MultipleChoiceNumberParameter`, `ChoiceNumber` (38 hits broad), `ChoiceNumberParameter` |
| **ConMas.iReporter.UserControls.dll** | Choice number WPF controls |
| **Cimtops.R2Cluster.dll** | `ChoiceNumber` referenced (4 hits) |
| **iReporterExcelAddInCommon.dll** | `ChoiceNumber` referenced (1 hit) |
| **LibConMas (property names)** | `ChoiceNumberExternal`, `ChoiceNumberReadOnly`, `ChoiceNumberClearOption`, `ChoiceNumberFontValue`, `ChoiceNumberFontSizeValue`, `ChoiceNumberWeightValue`, `ChoiceNumberColorValue`, `ChoiceNumberAlignValue`, `ChoiceNumberVerticalAlignValue`, `ChoiceNumberGroupValue` |

## Configuration Keys

### Verified Keys (from LibConMas property names)

| Key | Type | Default | Description | Evidence |
|-----|------|---------|-------------|----------|
| `Required` | bool | 0 | Whether selection is required | Common parameter |
| `ChoiceNumberExternal` | bool | 0 | External data source flag | LibConMas (3 hits) |
| `ChoiceNumberReadOnly` | bool | 0 | Read-only mode | LibConMas (3 hits) |
| `ChoiceNumberClearOption` | bool | 0 | Show clear/reset option | LibConMas (3 hits) |
| `ChoiceNumberGroupValue` | int | 0 | Option group number | LibConMas (3 hits) |
| `ChoiceNumberFontValue` | string | Arial | Font family | LibConMas (3 hits) |
| `ChoiceNumberFontSizeValue` | int | 11 | Font size | LibConMas (3 hits) |
| `ChoiceNumberWeightValue` | enum | Normal | Font weight | LibConMas (3 hits) |
| `ChoiceNumberColorValue` | string | 0,0,0 | Text color (R,G,B) | LibConMas (3 hits) |
| `ChoiceNumberAlignValue` | enum | Center | Text alignment | LibConMas (3 hits) |
| `ChoiceNumberVerticalAlignValue` | int | 2 | Vertical alignment | LibConMas (3 hits) |

### Common Parameters (shared with all field types)

| Key | Type | Default | Evidence |
|-----|------|---------|----------|
| `ReadOnly` | bool | 0 | ClusterParameter base |
| `Hidden` | bool | 0 | UserControls |
| `DefaultValue` | string | (empty) | Common parameter |

## Evidence Pattern

ChoiceNumber uses the same `*Value` suffix pattern as MultiSelect for property overrides:
```
ChoiceNumberFontValue       → Font
ChoiceNumberFontSizeValue  → FontSize
ChoiceNumberWeightValue   → Weight
ChoiceNumberColorValue    → Color
ChoiceNumberAlignValue    → Align
```

The `Value` suffix distinguishes runtime values from static definitions in the parameter class.

## Differences from MultiSelect

| Aspect | ChoiceNumber | MultiSelect |
|--------|-------------|-------------|
| cluster_type | `ChoiceNumber` | `MultiSelect` |
| Selection | Single (radio button group) | Multiple (checkboxes) |
| Property prefix | `ChoiceNumber*` | `MultiSelect*` |
| Extra params | `ChoiceNumberGroupValue`, `ChoiceNumberClearOption` | `MultiSelectLines`, `MultiSelectMaxValue` |

## Storage

- **Database:** `def_cluster.input_parameter` column
- **XML:** `<inputParameters>...</inputParameters>` within `<cluster>` element
- **Memory:** `MultipleChoiceNumberParameter` class

## Confidence Summary

| Item | Confidence |
|------|------------|
| ChoiceNumber as field type | ★★★★★ |
| MultipleChoiceNumberParameter class | ★★★★★ |
| 11 configuration properties | ★★★★☆ |
| Value-suffixed property pattern | ★★★★☆ |
| Difference from MultiSelect | ★★★☆☆ |
