# Field Configuration Reference (input_parameter)

## Serialization Format

**Format:** Semicolon-delimited `Key=Value` pairs.

```
Required=0;Lines=1;InputRestriction=None;MaxLength=0;Align=Center;Font=Arial;FontSize=11;Weight=Normal;Color=0,0,0;VerticalAlignment=2;DefaultFontSize=11
```

**Evidence:**
- `XmlGenerator.cs` (LegacyEngine) line 104: hardcoded example string
- `TabletPipeline.md` line 322-324: parsed from DB records
- `LibConMas.dll`: `InputParametersString`, `ParseParameterText`, `GetParametersArray` functions

**Delimiter:** `;` (semicolon) between pairs, `=` between key and value.

**Values:**
- `0`/`1` for boolean (false/true)
- Integer for numeric
- `R,G,B` for color (comma-separated 0-255)
- String for text values

## Master Parameter Table

### Common Parameters (ALL field types)

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | DLL + XML evidence |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | DLL (117 hits), defined in ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | DLL (33 hits in UserControls) |
| `DefaultValue` | string | (empty) | any | ★★★☆☆ | DLL (9 hits) |
| `Caption` | string | (empty) | any | ★★★☆☆ | DLL (10 hits) |
| `GroupName` | string | (empty) | any | ★☆☆☆☆ | DLL (1 hit) |

### KeyboardText

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `Lines` | int | 1 | 1+ | ★★★★★ | DLL (34 hits), in XML sample |
| `MaxLength` | int | 0 | 0=unlimited | ★★★★☆ | DLL (6 hits), in XML sample |
| `InputRestriction` | enum | None | None/Numeric/Alphabet | ★★★★☆ | DLL (7 hits), in XML sample |
| `Font` | string | Arial | font name | ★★★★☆ | DLL (21 hits), XML sample |
| `FontSize` | int | 11 | font size | ★★★★★ | DLL (57 hits), XML sample |
| `DefaultFontSize` | int | 11 | font size | ★★★☆☆ | DLL (3 hits), XML sample |
| `Weight` | enum | Normal | Normal/Bold | ★★★★☆ | DLL (23 hits), XML sample |
| `Color` | string | 0,0,0 | R,G,B | ★★★★☆ | DLL (82 hits), XML sample |
| `Align` | enum | Center | Left/Center/Right | ★★★★☆ | DLL (46 hits), XML sample |
| `VerticalAlignment` | int | 2 | 0=Top/1=Center/2=Bottom | ★★★★☆ | DLL (9 hits), XML sample |
| `Padding` | int | 0 | padding px | ★★★☆☆ | DLL (12 hits) |
| `Placeholder` | string | (empty) | placeholder text | ★☆☆☆☆ | LocalizableStrings.xml has "Placeholder" |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### InputNumeric

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `MinValue` | decimal | (none) | any | ★★★★★ | DLL (124 hits) |
| `MaxValue` | decimal | (none) | any | ★★★★★ | DLL (139 hits) |
| `Format` | string | (none) | number format | ★★★☆☆ | DLL (21 hits) |
| `DecimalPlaces` | int | 0 | 0+ | ★☆☆☆☆ | DLL (1 hit - `DecimalPlaces`) |
| `Font` | string | Arial | font name | ★★★★☆ | |
| `FontSize` | int | 11 | font size | ★★★★★ | |
| `Weight` | enum | Normal | Normal/Bold | ★★★★☆ | |
| `Color` | string | 0,0,0 | R,G,B | ★★★★☆ | |
| `Align` | enum | Center | Left/Center/Right | ★★★★☆ | |
| `VerticalAlignment` | int | 2 | 0=Top/1=Center/2=Bottom | ★★★★☆ | |
| `Padding` | int | 0 | padding px | ★★★☆☆ | |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** `InputNumeric.Error.Title`, `InputNumeric.Error.Minimum`, `InputNumeric.Error.Maximum` in LocalizableStrings.xml confirm validation messages.

### NumberHours

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `NumberHoursFormat` | string | (none) | format | ★★★★☆ | DLL (9 hits) |
| `NumberHoursDecimal` | int | 0 | decimal places | ★★★☆☆ | DLL (3 hits) |
| `NumberHoursMin` | decimal | (none) | min hours | ★★★☆☆ | DLL (3 hits / "NumberHoursMax") |
| `NumberHoursMax` | decimal | (none) | max hours | ★★★☆☆ | DLL (3 hits / "NumberHoursMin") |
| `NumberHoursAutoCalc` | bool | 0 | 0/1 | ★☆☆☆☆ | Hypothesis |
| `NumberHoursRoundUp` | bool | 0 | 0/1 | ★☆☆☆☆ | Hypothesis |
| `NumberHoursRoundDown` | bool | 0 | 0/1 | ★☆☆☆☆ | Hypothesis |
| `NumberHoursRoundMinute` | int | 0 | round to minute | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** `NumberHours.Error.Title` in LocalizableStrings.xml confirms error handling.

### ChoiceNumber

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `ChoiceNumberExternal` | bool | 0 | 0/1 | ★★★☆☆ | DLL (3 hits) |
| `ChoiceNumberReadOnly` | bool | 0 | 0/1 | ★★★☆☆ | DLL (3 hits) |
| `ChoiceNumberClearOption` | bool | 0 | 0/1 | ★★★☆☆ | DLL (3 hits) |
| `ChoiceNumberFontValue` | string | Arial | font | ★★★☆☆ | DLL (3 hits) |
| `ChoiceNumberFontSizeValue` | int | 11 | font size | ★★★☆☆ | DLL (3 hits) |
| `ChoiceNumberWeightValue` | enum | Normal | Normal/Bold | ★★★☆☆ | DLL (3 hits) |
| `ChoiceNumberColorValue` | string | 0,0,0 | R,G,B | ★★★☆☆ | DLL (3 hits) |
| `ChoiceNumberAlignValue` | enum | Center | Left/Center/Right | ★★★☆☆ | DLL (3 hits) |
| `ChoiceNumberVerticalAlignValue` | int | 2 | 0=Top/1=Center/2=Bottom | ★★★☆☆ | DLL (3 hits) |
| `ChoiceNumberGroupValue` | int | 0 | group number | ★★★☆☆ | DLL (3 hits) |
| `Font` | string | Arial | font name | ★★★★☆ | |
| `FontSize` | int | 11 | font size | ★★★★★ | |
| `Weight` | enum | Normal | Normal/Bold | ★★★★☆ | |
| `Color` | string | 0,0,0 | R,G,B | ★★★★☆ | |
| `Align` | enum | Center | Left/Center/Right | ★★★★☆ | |
| `VerticalAlignment` | int | 2 | 0=Top/1=Center/2=Bottom | ★★★★☆ | |
| `Padding` | int | 0 | padding px | ★★★☆☆ | |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### MultiSelect

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `MultiSelectExternal` | bool | 0 | 0/1 | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectReadOnly` | bool | 0 | 0/1 | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectValue` | string | (various) | selected values | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectDisplayValue` | string | (various) | display values | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectFontValue` | string | Arial | font | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectFontSizeValue` | int | 11 | font size | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectWeightValue` | enum | Normal | Normal/Bold | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectColorValue` | string | 0,0,0 | R,G,B | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectAlignValue` | enum | Center | Left/Center/Right | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectVerticalAlignValue` | int | 2 | 0=Top/1=Center/2=Bottom | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectPaddingValue` | int | 0 | padding px | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectMaxValue` | int | 0 | max selections | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectLines` | int | 0 | lines | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectUseKeyboardCheck` | bool | 0 | 0/1 | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectEditLabelCheck` | bool | 0 | 0/1 | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectInputRestrictionValue` | enum | None | None/Numeric | ★★★☆☆ | DLL (3 hits) |
| `MultiSelectProhibitedCharValue` | string | (empty) | prohibited chars | ★★★☆☆ | DLL (3 hits) |
| `Font` | string | Arial | font name | ★★★★☆ | |
| `FontSize` | int | 11 | font size | ★★★★★ | |
| `Weight` | enum | Normal | Normal/Bold | ★★★★☆ | |
| `Color` | string | 0,0,0 | R,G,B | ★★★★☆ | |
| `Align` | enum | Center | Left/Center/Right | ★★★★☆ | |
| `Padding` | int | 0 | padding px | ★★★☆☆ | |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### Calendar / CalendarDate

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `DateFormat` | string | (system) | date format | ★★★★★ | DLL (188 hits) |
| `TimeFormat` | string | (system) | time format | ★★★★☆ | DLL (64 hits) |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |
| `DefaultValue` | string | (empty) | date/time | ★★★☆☆ | |

**Note:** `CalendarDateForm`, `CalendarDateDialog` found in LibConMas.dll designer code.
`CalendarLibrary.Monday` in LocalizableStrings.xml confirms calendar day names.

### Date

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `DateFormat` | string | (system) | date format | ★★★★★ | DLL (188 hits) |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### Time

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `TimeFormat` | string | (system) | time format | ★★★★☆ | DLL (64 hits) |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### TimeCalculate

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `TimeFormat` | string | (system) | time format | ★★★★☆ | |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Note:** `TimeCalculateTextBox`, `TimeCalculateTimeForm` in UserControls.dll.

### Check

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `Caption` | string | (empty) | label text | ★★★☆☆ | |
| `CheckStyle` | enum | (default) | style | ★☆☆☆☆ | Hypothesis - `CheckClusterParameter` class exists |
| `CheckSize` | int | (default) | size | ★☆☆☆☆ | Hypothesis |
| `CheckColor` | string | (default) | R,G,B | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Note:** `CheckClusterParameter` class exists in LibConMas.dll.
`CheckMinCluster`, `CheckMaxCluster`, `CheckAllowMinCluster`, `CheckAllowMaxCluster` found in ConMasClient.exe for validation settings.

### Toggle

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `Caption` | string | (empty) | label text | ★★★☆☆ | |
| `ToggleStyle` | enum | (default) | style | ★☆☆☆☆ | Hypothesis - `Toggle` found 70 times in LibConMas.dll |
| `ToggleSize` | int | (default) | size | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### Image

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `ImageSize` | int/string | (default) | size | ★★★★☆ | DLL (44 hits) |
| `ImageFormat` | enum | (default) | JPG/PNG | ★★★★☆ | DLL (7 hits in LibConMas) |
| `ImageSource` | enum | (default) | Camera/Album/Both | ★★★☆☆ | DLL (2 hits) |
| `Resolution` | int | 0 | DPI | ★★★★☆ | DLL (53 hits in UserControls) |
| `Compression` | int | 0 | compression | ★★★☆☆ | DLL (8 hits in UserControls) |
| `ImageQuality` | int | 0 | quality % | ★☆☆☆☆ | Hypothesis |
| `CameraType` | enum | (default) | Front/Back | ★☆☆☆☆ | Hypothesis |
| `AutoCapture` | bool | 0 | 0/1 | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** `CameraControl`, `CameraDialog`, `PhotoLibrary`, `ShowCameraControl` in UserControls.dll.
`FailToCaptureImage` in LocalizableStrings.xml.

### FreeText / HandWriting

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `FreeTextMode` | enum | (default) | mode | ★☆☆☆☆ | Hypothesis |
| `FreeTextAutoInput` | bool | 0 | 0/1 | ★★★☆☆ | Settings detail HandWriting.AutoInput |
| `FreeTextAutoInputInterval` | int | 0 | ms | ★★★☆☆ | Settings detail HandWriting.AutoInputInterval |
| `FreeTextPadHeight` | int | (default) | px | ★★★☆☆ | Settings detail HandWriting.PadHeight |
| `FreeTextAreaHeight` | int | (default) | px | ★★★☆☆ | Settings detail HandWriting.AreaHeight |
| `FreeTextPadAlpha` | byte | (default) | 0-255 | ★★★☆☆ | Settings detail HandWriting.PadAlpha |
| `FreeTextSpacing` | int | (default) | px | ★★★☆☆ | Settings detail HandWriting.Spacing |
| `FreeTextArea` | bool | (default) | 0/1 | ★★★☆☆ | Settings.Drawing.Area |
| `FreeTextHeight` | int | (default) | px | ★★★☆☆ | Settings.Drawing.Height |
| `FreeTextTransparent` | bool | (default) | 0/1 | ★★★☆☆ | Settings.Drawing.Transparent |
| `FreeTextBaseHeight` | int | (default) | px | ★★★☆☆ | Settings.Drawing.BaseHeight |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** `HandWritingClusterParameter` in LibConMas.dll.
InkCanvas-based rendering in UserControls.dll (InkCanvas, StrokeCollection, DrawingAttributes).
Handwriting recognition engine: `AmiDnnMobile16k` (neural network).

### FreeDraw

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `FreeDrawMode` | enum | (default) | mode | ★☆☆☆☆ | Hypothesis - `FreeDrawClusterParameter` exists |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### Drawing / DrawingImage

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `DrawingMode` | enum | (default) | mode | ★☆☆☆☆ | Hypothesis - `DrawingImageClusterParameter` exists |
| `DrawingColor` | string | (default) | R,G,B | ★☆☆☆☆ | Hypothesis |
| `DrawingWidth` | int | (default) | px | ★☆☆☆☆ | Hypothesis |
| `DrawingStyle` | enum | (default) | style | ★☆☆☆☆ | Hypothesis |
| `DrawingEraserSize` | int | (default) | px | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### Barcode

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `BarcodeFormat` | enum | (default) | format | ★★★☆☆ | DLL (44 hits "Barcode") |
| `ScanMode` | enum | (default) | Single/Continuous | ★★★★☆ | DLL (9 hits) |
| `Timeout` | int | 0 | ms | ★★★★☆ | DLL (10 hits) |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Library:** zxing.dll for barcode decoding.

### QRCode

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `QRCodeFormat` | enum | (default) | format | ★★★☆☆ | `QRCodeClusterParameter` exists |
| `ScanMode` | enum | (default) | Single/Continuous | ★★★★☆ | |
| `Timeout` | int | 0 | ms | ★★★★☆ | |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### EdgeOCR

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `EdgeOCRMode` | enum | (default) | mode | ★☆☆☆☆ | `EdgeOCRClusterParameter` exists (52 hits) |
| `EdgeOCRFormat` | enum | (default) | format | ★☆☆☆☆ | Hypothesis |
| `EdgeOCRLanguage` | string | (default) | language | ★★★★☆ | DLL (5 hits "Language") |
| `EdgeOCRRegion` | string | (default) | region | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** `EdgeOCRCluster`, `EdgeOCRScanSettingControl`, `EdgeOCROutputCluster` in ConMasClient.exe.
`EdgeOCRMethods` (9 hits) in LibConMas.dll.

### Scandit

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `ScanditMode` | enum | (default) | mode | ★★★★☆ | DLL (4 hits) |
| `ScanditFormat` | enum | (default) | format | ★☆☆☆☆ | `ScanditClusterParameter` exists |
| `Language` | string | (default) | language | ★★★★☆ | DLL (5 hits) |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### GPS / Location

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `GPSUpdateInterval` | int | (default) | ms | ★☆☆☆☆ | Hypothesis - `GpsClusterParameter` exists |
| `GPSDistanceFilter` | int | (default) | meters | ★★★★☆ | LocalizableStrings.xml: 10m/100m/1000m |
| `GPSAccuracy` | int | (default) | meters | ★☆☆☆☆ | Hypothesis |
| `GPSAutoStop` | bool | 0 | 0/1 | ★☆☆☆☆ | Hypothesis |
| `Timeout` | int | 0 | ms | ★★★★☆ | |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### Action

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `ActionType` | enum | (default) | Command/Output/SetValue | ★★★☆☆ | DLL (1 hit `ActionType`) |
| `ActionValue` | string | (empty) | action target | ★★★★☆ | DLL (6 hits) |
| `ActionLabel` | string | (empty) | button label | ★☆☆☆☆ | Hypothesis |
| `ActionConfirm` | bool | 0 | 0/1 | ★☆☆☆☆ | Hypothesis |
| `ActionConfirmText` | string | (empty) | confirm text | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** `Cluster.Action.DisplayValue.CommandDone`, `Cluster.Action.DisplayValue.OutputDone`, `Cluster.Action.DisplayValue.SetValueDone` in LocalizableStrings.xml.
`LblActionButton`, `TextModeActionButton` in ConMasClient.exe.

### SelectMaster

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `MasterType` | enum | (default) | type | ★★★☆☆ | DLL (3 hits) |
| `MasterValue` | string | (empty) | value | ★★★☆☆ | DLL (3 hits) |
| `MasterKey` | string | (empty) | key field | ★★★★☆ | DLL (7 hits) |
| `MasterSource` | string | (empty) | source | ★☆☆☆☆ | Hypothesis |
| `MasterDisplay` | string | (empty) | display field | ★☆☆☆☆ | Hypothesis |
| `MasterFilter` | string | (empty) | filter | ★☆☆☆☆ | Hypothesis |
| `MasterCache` | bool | 1 | 0/1 | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** `MasterInfoService` in LibConMas.dll.

### LoginUser

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `ReadOnly` | bool | 1 | 0/1 | ★★★★★ | System-generated, always read-only |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** System-generated field populated with current login user name/user ID.
`LoginUserDialog` in ConMasClient.exe. `LoginUserId` in LibConMas.dll.

### Issuer

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `ReadOnly` | bool | 1 | 0/1 | ★★★★★ | System-generated, always read-only |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** System-generated field populated with the issuer of the document.
`Issuer` found in LibConMas.dll and iReporterExcelAddInCommon.dll.

### Approve

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `ApproveRequired` | bool | 1 | 0/1 | ★★★★☆ | DLL (6 hits) |
| `ApproveType` | enum | (default) | type | ★☆☆☆☆ | Hypothesis |
| `ApproveSign` | bool | 1 | 0/1 | ★★★★☆ | DLL (5 hits) |
| `ApproveDateVisible` | bool | 1 | 0/1 | ★★★★☆ | `ApproveCreateDateValue` in UserControls |
| `ApproveCommentVisible` | bool | 1 | 0/1 | ★★★★☆ | `ApproveCreateCommentValue` in UserControls |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** `SimpleApproveDialog` in UserControls.dll.
`Approve.Type.Editting`, `Approve.Type.Unedited` in LocalizableStrings.xml.
`ApproveCreateDateValue`, `ApproveUserIdValue`, `ApproveUserNameValue` show what's stored.
`ApproveButtonTitle`, `ApproveCreateCommentTitle`, `ApproveCreateDateTitle` in UserControls.dll.

### Inspect

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `InspectType` | enum | (default) | type | ★☆☆☆☆ | Hypothesis |
| `InspectResult` | string | (empty) | result options | ★☆☆☆☆ | Hypothesis |
| `InspectLabel` | string | (empty) | label | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### Issue

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `IssueType` | enum | (default) | type | ★☆☆☆☆ | Hypothesis |
| `IssueAutoGenerate` | bool | 1 | 0/1 | ★☆☆☆☆ | Hypothesis |
| `IssuePrefix` | string | (empty) | prefix | ★☆☆☆☆ | Hypothesis |
| `IssueNumber` | int | 0 | number | ★☆☆☆☆ | Hypothesis |
| `IssueDigits` | int | 0 | digit count | ★☆☆☆☆ | Hypothesis |
| `IssueDate` | bool | 1 | 0/1 | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

### PinNo

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `PinImage` | string | (empty) | image ref | ★★★★☆ | DLL (6 hits) |
| `PinColor` | string | (default) | R,G,B | ★★★★☆ | DLL (4 hits) |
| `PinSize` | int | (default) | px | ★☆☆☆☆ | Hypothesis |
| `PinLabel` | string | (empty) | label | ★☆☆☆☆ | Hypothesis |
| `PinLabelPosition` | enum | (default) | position | ★☆☆☆☆ | Hypothesis |
| `PinShape` | enum | (default) | shape | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** `DrawingPinNo` in UserControls.dll. `PinItemTableNo`, `PinItemCluster` in LibConMas.dll.
`MovePinWindow` in UserControls.dll. `PinCooperationSelectCluster` in ConMasClient.exe.

### AudioRecording

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `AudioRecordingFormat` | enum | (default) | WAV/MP3 | ★★★☆☆ | `AudioRecordingFormat` in UserControls |
| `AudioMode` | enum | (default) | mode | ★☆☆☆☆ | Hypothesis |
| `AudioQuality` | enum | (default) | quality | ★☆☆☆☆ | Hypothesis |
| `AudioMaxDuration` | int | 0 | seconds | ★☆☆☆☆ | Hypothesis |
| `AudioAutoStart` | bool | 0 | 0/1 | ★☆☆☆☆ | Hypothesis |
| `AudioAutoStop` | bool | 0 | 0/1 | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 0 | 0/1 | ★★★★★ | |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** `AudioRecordingParameter` in UserControls.dll.
`AudioRecordingPage`, `AudioRecordingPreparationPage`, `AudioReproducingPage` in UserControls.
`RecordWindow`, `AudioRecordingWindow`, `CustomRecordWindow` for recording UI.
NAudio.dll for audio capture.
`RecordingStartImage`, `RecordingButtonImage`, `RecordingSeconds` in UserControls.

### Calculation

| Parameter | Type | Default | Values | Verified | Notes |
|-----------|------|---------|--------|----------|-------|
| `Required` | bool | 0 | 0/1 | ★★★★☆ | |
| `CalculationFormula` | string | (empty) | formula | ★☆☆☆☆ | Hypothesis - `Calculation` found 3 times |
| `CalculationDecimal` | int | 0 | decimal places | ★☆☆☆☆ | Hypothesis |
| `CalculationFormat` | string | (empty) | format | ★☆☆☆☆ | Hypothesis |
| `ReadOnly` | bool | 1 | 0/1 | ★★★★★ | Calculated fields are auto-computed |
| `Hidden` | bool | 0 | 0/1 | ★★★★☆ | |

**Runtime:** `CalculateDateForm`, `TimeCalForm` in ConMasClient.exe.
`CalculateMaxCluster`, `CalculateMinCluster`, `CalculateAllowMaxCluster`, `CalculateAllowMinCluster` validation.

## System-Generated Fields

These field types are automatically populated by the system and typically have no configurable input_parameter:

| Field Type | Population Rule | Evidence |
|------------|----------------|----------|
| `LatestUpdateDate` | System date/time of last update | `LatestUpdateDateClusterParameter` in LibConMas.dll |
| `RegistrationDate` | System date/time of creation | `RegistrationDateClusterParameter` in LibConMas.dll |
| `ExpiryDate` | Expiry date from form settings | `ExpiryDateForm` in ConMasClient.exe |
| `LoginUser` | Current user name/ID | `LoginUserClusterParameter` in LibConMas.dll |
| `Issuer` | Issuing user name | `Issuer` in LibConMas.dll / iReporterExcelAddInCommon.dll |

## Confidence Scale

| Rating | Meaning |
|--------|---------|
| ★★★★★ | Verified in XML + DLL + Database |
| ★★★★☆ | Verified in DLL + Database (or multiple DLLs) |
| ★★★☆☆ | Found in DLL with strong evidence |
| ★★☆☆☆ | Found in DLL but limited evidence |
| ★☆☆☆☆ | Hypothesis based on naming patterns |

## Evidence Sources

| Source | File | Relevance |
|--------|------|-----------|
| LibConMas.dll | `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\bin\LibConMas.dll` | 60+ ClusterParameter classes, serialization functions |
| ConMas.iReporter.UserControls.dll | `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas i-Reporter for Windows\ConMas.iReporter.UserControls.dll` | WPF controls, parameter dialogs |
| ConMasClient.exe | `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\bin\ConMasClient.exe` | Designer property pages, setting windows |
| Cimtops.R2Cluster.dll | `C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn\Cimtops.R2Cluster.dll` | Cluster management, type constants |
| iReporterExcelAddInCommon.dll | `C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn\iReporterExcelAddInCommon.dll` | ClusterInfo, shared logic |
| XmlGenerator.cs (LegacyEngine) | `ExcelAPI/ExcelAPI/LegacyEngine/PublishEngine/XmlGenerator.cs` | Hardcoded inputParameters sample |
| LocalizableStrings.xml | `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas i-Reporter for Windows\LocalizableStrings.xml` | UI strings revealing parameter names |
| DefinitionBiz.xml | `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas i-Reporter for Windows\xml\DefinitionBiz.xml` | Database write queries |
