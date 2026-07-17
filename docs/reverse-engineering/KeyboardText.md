# KeyboardText — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `KeyboardText` |
| XML `<type>` | `KeyboardText` |
| Comment line 1 | `KeyboardText` |
| Parameter class | `TextClusterParameter` / `KeyboardTextClusterParameter` |
| DLL class count | `TextClusterParameter` (2 hits), `KeyboardTextClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `TextClusterParameter` class, `InputParameterDefault`, `ParseParameterText` |
| **ConMas.iReporter.UserControls.dll** | WPF text controls, keyboard dialogs |
| **XmlGenerator.cs** (LegacyEngine) | Hardcoded sample: `Required=0;Lines=1;InputRestriction=None;MaxLength=0;Align=Center;Font=Arial;FontSize=11;Weight=Normal;Color=0,0,0;VerticalAlignment=2;DefaultFontSize=11` |
| **DB Templates 546, 547** | All 11 clusters are KeyboardText type (6+5) |
| **LocalizableStrings.xml** | `KeyboardText.Label.Font`, `KeyboardText.Label.Size`, `KeyboardText.Label.Overflow`, `KeyboardText.Label.Done`, `KeyboardText.Title.ChooseColor` |
| **ConMasClient.exe** | `TxtDateForm`, `TxtTimeForm` (text-based date/time input forms) |

## Configuration Keys

### Verified Keys (from XML + DLL + Database)

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether input is mandatory | XML sample, common parameter |
| `Lines` | int | 1 | 1+ | Number of text lines (1=single, >1=multiline) | XML sample (Lines=1) |
| `InputRestriction` | enum | None | None/Numeric/Alphabet | Restrict allowed characters | XML sample (InputRestriction=None) |
| `MaxLength` | int | 0 | 0=unlimited | Maximum character count | XML sample (MaxLength=0) |
| `Align` | enum | Center | Left/Center/Right | Horizontal text alignment | XML sample (Align=Center) |
| `Font` | string | Arial | Font name | Font family name | XML sample (Font=Arial) |
| `FontSize` | int | 11 | Font size | Font size in points | XML sample (FontSize=11) |
| `DefaultFontSize` | int | 11 | Font size | Default font size when empty | XML sample (DefaultFontSize=11) |
| `Weight` | enum | Normal | Normal/Bold | Font weight | XML sample (Weight=Normal) |
| `Color` | string | 0,0,0 | R,G,B | Text color as comma-separated RGB | XML sample (Color=0,0,0) |
| `VerticalAlignment` | int | 2 | 0=Top/1=Center/2=Bottom | Vertical text alignment | XML sample (VerticalAlignment=2) |

### Keys from DLL (strong evidence, not in XML sample)

| Key | Type | Default | Values | Evidence |
|-----|------|---------|--------|----------|
| `ReadOnly` | bool | 0 | 0/1 | LibConMas (117 hits), ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | UserControls (33 hits) |
| `Padding` | int | 0 | Padding px | LibConMas (12 hits), UserControls (6 hits) |
| `Placeholder` | string | (empty) | Placeholder text | LocalizableStrings.xml `Placeholder` entries |
| `DefaultValue` | string | (empty) | Any text | LibConMas (9 hits) |

### Hypothetical Keys (from naming patterns only)

| Key | Type | Default | Description |
|-----|------|---------|-------------|
| `TextAlign` | enum | Center | Alternative key for Align? |
| `TextColor` | string | 0,0,0 | Alternative key for Color? |
| `TextVerticalAlign` | int | 2 | Alternative key for VerticalAlignment? |
| `GroupName` | string | (empty) | Group/fieldset name |
| `Caption` | string | (empty) | Display label text |

## Default inputParameters String

```
Required=0;Lines=1;InputRestriction=None;MaxLength=0;Align=Center;Font=Arial;FontSize=11;Weight=Normal;Color=0,0,0;VerticalAlignment=2;DefaultFontSize=11
```

**Source:** `XmlGenerator.cs` (LegacyEngine/PublishEngine) line 104-106.

## Designer UI Evidence

From `LocalizableStrings.xml`:
- `KeyboardText.Label.Font` — Font selector
- `KeyboardText.Label.Size` — Font size selector
- `KeyboardText.Label.Overflow` — Overflow behavior
- `KeyboardText.Label.Done` — Done button
- `KeyboardText.Title.ChooseColor` — Color picker dialog title

## Runtime Behavior

The legacy runtime creates a WPF `TextBox` (single-line) or `TextBox` with `AcceptsReturn=True` (multiline) based on the `Lines` parameter.

Common WPF properties applied from input_parameter:
- `TextBox.TextAlignment` ← `Align`
- `TextBox.FontFamily` ← `Font`
- `TextBox.FontSize` ← `FontSize`
- `TextBox.FontWeight` ← `Weight`
- `TextBox.Foreground` ← `Color`
- `TextBox.MaxLength` ← `MaxLength`
- `TextBox.AcceptsReturn` ← `Lines > 1`
- `TextBox.VerticalContentAlignment` ← `VerticalAlignment`
- `TextBox.IsReadOnly` ← `ReadOnly`
- `TextBox.IsEnabled` ← `!ReadOnly` (alternative)

## Comment Format (Excel Add-In)

```
Line 0: Field Name (e.g., "samples")
Line 1: KeyboardText
Line 2: Cluster Index (sort order, optional)
Line 3+: (input parameters not stored in comment — set via InputForm2 dialog)
```

## Storage

- **Database:** `def_cluster.input_parameter` column (text, semicolon-delimited)
- **XML:** `<inputParameters>...</inputParameters>` within `<cluster>` element
- **Memory:** `TextClusterParameter` class in LibConMas.dll

## Confidence Summary

| Item | Confidence |
|------|------------|
| Serialization format (key=value;) | ★★★★★ |
| 11 core parameters (Required..DefaultFontSize) | ★★★★★ |
| 4 additional parameters (ReadOnly..DefaultValue) | ★★★★☆ |
| 1 hypothetical parameters | ★☆☆☆☆ |
| WPF TextBox as runtime control | ★★★★★ |
