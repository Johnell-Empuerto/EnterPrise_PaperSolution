# FreeText — Legacy Field Configuration (HandWriting / FreeText)

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `FreeText` |
| XML `<type>` | `FreeText` |
| Comment line 1 | `FreeText` |
| Parameter class | `FreeTextClusterParameter` / `HandwritingClusterParameter` |
| DLL class count | `FreeTextClusterParameter` (1 hit), `HandwritingClusterParameter` (1 hit) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `FreeTextClusterParameter`, `HandwritingClusterParameter`, `FreeText` (29 hits broad) |
| **ConMas.iReporter.UserControls.dll** | InkCanvas, InkPresenter, StrokeCollection, StylusPoint, DrawingAttributes |
| **LocalizableStrings.xml** | `MainMenu.Title.HandWriting`, `Settings.Menu.HandWritingPad`, `Settings.Detail.HandWriting.*` (8 settings) |
| **ClusterImageBiz.xml** | `cluster_type IN ('Image','FreeText')` — FreeText stores image data |
| **Res/systems/AmiDnnMobile16k** | Handwriting recognition engine (neural network) |
| **Settings.Drawing.* (LocalizableStrings)** | `Settings.Drawing.Area`, `Settings.Drawing.Height`, `Settings.Drawing.Transparent`, `Settings.Drawing.BaseHeight` |

## Configuration Keys

### Verified Keys (from LibConMas property names + LocalizableStrings)

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether input is required | Common parameter |
| `FreeTextAutoInput` | bool | 0 | 0/1 | Enable auto-input mode | Settings.Detail.HandWriting.AutoInput |
| `FreeTextAutoInputInterval` | int | (default) | ms | Auto-input interval | Settings.Detail.HandWriting.AutoInputInterval |
| `FreeTextPadHeight` | int | (default) | px | Handwriting pad height | Settings.Detail.HandWriting.PadHeight |
| `FreeTextAreaHeight` | int | (default) | px | Handwriting area height | Settings.Detail.HandWriting.AreaHeight |
| `FreeTextPadAlpha` | byte | (default) | 0-255 | Pad transparency (alpha) | Settings.Detail.HandWriting.PadAlpha |
| `FreeTextSpacing` | int | (default) | px | Character spacing | Settings.Detail.HandWriting.Spacing |
| `FreeTextAutoScroll` | bool | (default) | 0/1 | Auto scroll during input | Settings.Detail.HandWriting.AutoScroll |
| `FreeTextScrollRange` | int | (default) | px | Scroll range | Settings.Detail.HandWriting.ScrollRange |
| `FreeTextArea` | bool | (default) | 0/1 | Show writing area | Settings.Drawing.Area |
| `FreeTextHeight` | int | (default) | px | Writing area height | Settings.Drawing.Height |
| `FreeTextTransparent` | bool | (default) | 0/1 | Transparent background | Settings.Drawing.Transparent |
| `FreeTextBaseHeight` | int | (default) | px | Base line height | Settings.Drawing.BaseHeight |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

## Runtime Technology

### WPF InkCanvas
The runtime uses WPF `InkCanvas` for handwriting input:
- `InkCanvas` — Main drawing surface
- `StrokeCollection` — Collection of ink strokes
- `DrawingAttributes` — Pen attributes (color, width, style)
- `StylusPointCollection` — Stylus input points
- `InkCanvasEditingMode` — Editing mode (Ink/Select/Erase)

### Handwriting Recognition
Handwriting recognition engine at `res/systems/AmiDnnMobile16k/`:
- Neural network-based recognizer
- `recognizer.xc` — Recognition model file
- Recognizes handwritten Japanese and possibly Latin characters

## Data Storage

Like `Image`, `FreeText` stores its data in `rep_cluster.image_file` as base64-encoded image data.

From `ClusterImageBiz.xml`:
```sql
UPDATE rep_cluster SET image_file = :img_base
WHERE cluster_type IN ('Image','FreeText')
```

## Settings (from LocalizableStrings.xml)

### HandWriting Settings (`Settings.Detail.HandWriting.*`)
| Setting ID | UI String (EN) | UI String (JA) |
|-----------|----------------|----------------|
| AutoScroll | Auto Scroll | 自動スクロール |
| ScrollRange | Scroll Range | スクロール範囲 |
| AutoInput | Auto Input | 自動入力 |
| AutoInputInterval | Auto Input Interval | 自動入力間隔 |
| PadHeight | Pad Height | パッド高さ |
| AreaHeight | Area Height | エリア高さ |
| PadAlpha | Pad Alpha | パッド透明度 |
| Spacing | Spacing | 文字間隔 |

### Drawing Settings (`Settings.Drawing.*`)
| Setting ID | UI String (EN) | UI String (JA) |
|-----------|----------------|----------------|
| Area | Drawing Area | 描画エリア |
| Height | Drawing Height | 描画高さ |
| Transparent | Transparent Drawing | 透明描画 |
| BaseHeight | Base Height | ベース高さ |

## Differences from FreeDraw

| Aspect | FreeText | FreeDraw |
|--------|----------|----------|
| cluster_type | `FreeText` | `FreeDraw` |
| Use case | Handwriting input (characters) | Free-form drawing |
| Recognition | AmiDnnMobile16k handwriting recognition | No recognition |
| Storage | Base64 image in rep_cluster.image_file | Likely same |
| Parameters | Text/auto-input focused | Drawing style focused |

## Confidence Summary

| Item | Confidence |
|------|------------|
| FreeText as field type | ★★★★★ |
| InkCanvas rendering | ★★★★★ |
| Handwriting recognition engine | ★★★★★ |
| 8 HandWriting settings params | ★★★★☆ |
| 4 Drawing settings params | ★★★☆☆ |
| Data stored as base64 image | ★★★★★ |
