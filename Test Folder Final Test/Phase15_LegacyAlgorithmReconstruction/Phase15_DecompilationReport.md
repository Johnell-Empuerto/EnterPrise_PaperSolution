# Legacy Assembly Decompilation Report

**Generated:** 2026-07-12 20:25:47

## Overview

This report documents the legacy PaperLess application assemblies found at:

- `C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn\`
- `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator\`
- `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\`
- `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas i-Reporter for Windows\`

---

## iReporterExcelAddIn — Cimtops.Excel

**File:** `C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn\Cimtops.Excel.dll`
**Size:** 24,576 bytes

**Architecture:** x86 (32-bit)
**Managed (.NET) assembly:** YES
**CLR Header:** RVA=0x00002008, Size=0x00000048

### Types

| Type | Kind | Public Methods | Public Fields |
|------|------|---------------|--------------|
| `Cimtops.Excel.ActionTypes` | Abstract Class | 1 | 15 |
| `Cimtops.Excel.Book` | Class | 0 | 0 |
| `Cimtops.Excel.BookFactory` | Class | 1 | 0 |
| `Cimtops.Excel.BorderQ` | Class | 2 | 0 |
| `Cimtops.Excel.BorderQ+<>c` | Class | 0 | 2 |
| `Cimtops.Excel.BorderQ+<>c__DisplayClass13_0` | Class | 0 | 2 |
| `Cimtops.Excel.Cell` | Class | 0 | 0 |
| `Cimtops.Excel.Cell+<>c` | Class | 0 | 3 |
| `Cimtops.Excel.Cell+<>o__24` | Abstract Class | 0 | 6 |
| `Cimtops.Excel.Cell+<>o__25` | Abstract Class | 0 | 1 |
| `Cimtops.Excel.CellBorder` | Class | 2 | 0 |
| `Cimtops.Excel.CellModify` | Class | 0 | 2 |
| `Cimtops.Excel.CellRange` | Class | 4 | 0 |
| `Cimtops.Excel.Col` | Class | 0 | 0 |
| `Cimtops.Excel.DeviceType` | Enum | 0 | 4 |
| `Cimtops.Excel.Direction` | Enum | 0 | 6 |
| `Cimtops.Excel.Row` | Class | 2 | 0 |
| `Cimtops.Excel.Row+<>c__DisplayClass11_0` | Class | 0 | 1 |
| `Cimtops.Excel.Sheet` | Class | 5 | 0 |
| `Cimtops.Excel.Sheet+<>c` | Class | 0 | 3 |
| `Cimtops.Excel.Sheet+<>c__DisplayClass25_0` | Class | 0 | 1 |
| `Cimtops.Excel.Sheet+<>c__DisplayClass26_0` | Class | 0 | 1 |
| `Cimtops.Excel.Sheet+<>c__DisplayClass26_1` | Class | 0 | 2 |
| `Cimtops.Excel.Sheet+<>o__26` | Abstract Class | 0 | 3 |
| `Cimtops.Excel.XLRangeUtil` | Abstract Class | 3 | 0 |
| `Microsoft.Office.Interop.Excel._Workbook` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel._Worksheet` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.Border` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.Borders` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.Comment` | Interface | 1 | 0 |
| `Microsoft.Office.Interop.Excel.DocEvents` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.DocEvents_Event` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.Interior` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.Pages` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.PageSetup` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.Range` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.Sheets` | Interface | 1 | 0 |
| `Microsoft.Office.Interop.Excel.Workbook` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.WorkbookEvents` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.WorkbookEvents_Event` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.Worksheet` | Interface | 0 | 0 |
| `Microsoft.Office.Interop.Excel.XlBordersIndex` | Enum | 0 | 9 |
| `Microsoft.Office.Interop.Excel.XlLineStyle` | Enum | 0 | 9 |

### Detailed Methods (Coordinate-Related Types)

#### `Cimtops.Excel.ActionTypes`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `GetDevices` | `public static` | `DeviceType[]` | `String actionType` |

#### `Cimtops.Excel.Book`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.Excel.BookFactory`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `Create` | `public static` | `Book` | `Workbook book, Exception& ex` |
| `Create` | `public static` | `Sheet` | `Workbook book, Worksheet sheet, Exception& ex` |

#### `Cimtops.Excel.BorderQ`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<Exists>b__11_0` | `private` | `Boolean` | `Direction n` |
| `<get_IsEncloded>b__10_0` | `private` | `Boolean` | `CellBorder b` |
| `Exists` | `public` | `Boolean` | `Direction[] index` |
| `Exists` | `private` | `Boolean` | `CellBorder border, Boolean onlyReal` |
| `HasAll` | `public` | `Boolean` | `Direction[] index` |
| `HasAll` | `public` | `Boolean` | `Boolean onlyReal, Direction[] index` |

#### `Cimtops.Excel.BorderQ+<>c`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<.ctor>b__8_0` | `internal` | `CellBorder` | `Int32 n` |

#### `Cimtops.Excel.BorderQ+<>c__DisplayClass13_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<HasAll>b__0` | `internal` | `Boolean` | `Direction n` |

#### `Cimtops.Excel.Cell`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<get_Height>b__21_0` | `private` | `Boolean` | `Row r` |
| `<get_Width>b__19_0` | `private` | `Boolean` | `Col c` |
| `C2I` | `private` | `Int32` | `Object color` |
| `ToStyle` | `private` | `XlLineStyle` | `Borders borders, XlBordersIndex dir` |

#### `Cimtops.Excel.Cell+<>c`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<get_Height>b__21_1` | `internal` | `Double` | `Row r` |
| `<get_Width>b__19_1` | `internal` | `Double` | `Col c` |

#### `Cimtops.Excel.Cell+<>o__24`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.Excel.Cell+<>o__25`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.Excel.CellBorder`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `SetStyle` | `public` | `Void` | `XlLineStyle style` |
| `ToString` | `public` | `String` | `` |

#### `Cimtops.Excel.CellModify`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.Excel.CellRange`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `Contains` | `public` | `Boolean` | `Int32 row, Int32 col` |
| `IsOverlap` | `public` | `Boolean` | `CellRange other` |
| `ToString` | `public` | `String` | `` |
| `ToText` | `public` | `String` | `` |

#### `Cimtops.Excel.Col`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.Excel.DeviceType`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.Excel.Direction`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.Excel.Row`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `Add` | `public` | `Void` | `Cell cell` |
| `Find` | `public` | `Cell` | `Int32 x` |

#### `Cimtops.Excel.Row+<>c__DisplayClass11_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<Find>b__0` | `internal` | `Boolean` | `Cell c` |

#### `Cimtops.Excel.Sheet`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `DownOf` | `public` | `Cell` | `Cell cell` |
| `GetCell` | `public` | `Cell` | `Int32 row, Int32 col` |
| `LeftOf` | `public` | `Cell` | `Cell cell` |
| `RightOf` | `public` | `Cell` | `Cell cell` |
| `UpOf` | `public` | `Cell` | `Cell cell` |

#### `Cimtops.Excel.Sheet+<>c`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<get_Height>b__20_0` | `internal` | `Double` | `Row c` |
| `<get_Width>b__18_0` | `internal` | `Double` | `Col c` |

#### `Cimtops.Excel.Sheet+<>c__DisplayClass25_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<GetCell>b__0` | `internal` | `Boolean` | `Row r` |

#### `Cimtops.Excel.Sheet+<>c__DisplayClass26_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.Excel.Sheet+<>c__DisplayClass26_1`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<.ctor>b__0` | `internal` | `Boolean` | `Cell c` |

#### `Cimtops.Excel.Sheet+<>o__26`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.Excel.XLRangeUtil`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `AdjustRange` | `public static` | `String` | `String area` |
| `Num2Alpha` | `private static` | `String` | `Int32 number` |
| `ToAddressName` | `public static` | `String` | `Int32 rowNum, Int32 colNum, Int32 firstNum` |
| `ToRowCol` | `public static` | `Tuple`2` | `String area` |

#### `Microsoft.Office.Interop.Excel._Workbook`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel._Worksheet`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.Border`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.Borders`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.Comment`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `Text` | `public` | `String` | `Object Text, Object Start, Object Overwrite` |

#### `Microsoft.Office.Interop.Excel.DocEvents`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.DocEvents_Event`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.Interior`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.Pages`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.PageSetup`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.Range`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.Sheets`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `GetEnumerator` | `public` | `IEnumerator` | `` |

#### `Microsoft.Office.Interop.Excel.Workbook`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.WorkbookEvents`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.WorkbookEvents_Event`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.Worksheet`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.XlBordersIndex`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Microsoft.Office.Interop.Excel.XlLineStyle`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

## iReporterExcelAddIn — Cimtops.R2Cluster

**File:** `C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn\Cimtops.R2Cluster.dll`
**Size:** 42,496 bytes

**Architecture:** x86 (32-bit)
**Managed (.NET) assembly:** YES
**CLR Header:** RVA=0x00002008, Size=0x00000048

### Types

| Type | Kind | Public Methods | Public Fields |
|------|------|---------------|--------------|
| `Cimtops.R2Cluster.Caption` | Struct | 4 | 0 |
| `Cimtops.R2Cluster.CaptionPriority` | Enum | 0 | 5 |
| `Cimtops.R2Cluster.Cluster` | Class | 7 | 0 |
| `Cimtops.R2Cluster.Cluster+<>c` | Class | 0 | 3 |
| `Cimtops.R2Cluster.Cluster+<>c__DisplayClass46_0` | Class | 0 | 2 |
| `Cimtops.R2Cluster.Cluster+<>c__DisplayClass53_0` | Class | 0 | 1 |
| `Cimtops.R2Cluster.Cluster+<>c__DisplayClass54_0` | Class | 0 | 1 |
| `Cimtops.R2Cluster.Cluster+<>c__DisplayClass54_1` | Class | 0 | 2 |
| `Cimtops.R2Cluster.Cluster+<>c__DisplayClass54_2` | Class | 0 | 2 |
| `Cimtops.R2Cluster.Decoder` | Class | 7 | 0 |
| `Cimtops.R2Cluster.Decoder+<>c` | Class | 0 | 3 |
| `Cimtops.R2Cluster.Decoder+<>c__DisplayClass24_0` | Class | 0 | 1 |
| `Cimtops.R2Cluster.Decoder+<>c__DisplayClass25_0` | Struct | 0 | 2 |
| `Cimtops.R2Cluster.Decoder+<>c__DisplayClass33_0` | Class | 0 | 1 |
| `Cimtops.R2Cluster.Decoder+<>c__DisplayClass35_0` | Struct | 0 | 4 |
| `Cimtops.R2Cluster.Decoder+<>c__DisplayClass36_0` | Class | 0 | 3 |
| `Cimtops.R2Cluster.Decoder+<>c__DisplayClass36_1` | Class | 0 | 2 |
| `Cimtops.R2Cluster.Decoder+<>c__DisplayClass36_2` | Class | 0 | 2 |
| `Cimtops.R2Cluster.Decoder+<SplitLine>d__29` | Class | 0 | 1 |
| `Cimtops.R2Cluster.Decoder+AIInfo` | Struct | 0 | 0 |
| `Cimtops.R2Cluster.Decoder+AutoJudgeResult` | Class | 0 | 0 |
| `Cimtops.R2Cluster.Decoder+ClusterInfo` | Struct | 1 | 7 |
| `Cimtops.R2Cluster.Global` | Abstract Class | 0 | 0 |
| `Cimtops.R2Cluster.JudgeResult` | Class | 6 | 1 |
| `Cimtops.R2Cluster.JudgeResult+<>c__DisplayClass25_0` | Class | 0 | 4 |
| `Cimtops.R2Cluster.JudgeResult+<>c__DisplayClass4_0` | Class | 0 | 1 |
| `Cimtops.R2Cluster.JudgeResult+Judement` | Class | 2 | 0 |
| `Cimtops.R2Cluster.JudgeResult+Judement+<>c__DisplayClass32_0` | Class | 0 | 2 |
| `Cimtops.R2Cluster.TitleInfo` | Class | 0 | 2 |
| `Cimtops.R2Cluster.Util` | Abstract Class | 2 | 0 |
| `Microsoft.CodeAnalysis.EmbeddedAttribute` | Class | 0 | 0 |
| `System.Runtime.CompilerServices.IsReadOnlyAttribute` | Class | 0 | 0 |

### Detailed Methods (Coordinate-Related Types)

#### `Cimtops.R2Cluster.Caption`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `AdjustText` | `private` | `Void` | `` |
| `Contains` | `public` | `Boolean` | `CellRange other` |
| `GetText` | `public` | `String` | `` |
| `IsPriorThan` | `public` | `Boolean` | `Caption other` |
| `SetFailed` | `public` | `Void` | `String reason` |

#### `Cimtops.R2Cluster.CaptionPriority`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.R2Cluster.Cluster`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `AddAround` | `internal` | `Void` | `String text` |
| `AddCaption` | `private` | `Void` | `Caption cap` |
| `AdustCaption` | `internal` | `Void` | `Sheet sheet, List`1 clusters` |
| `Contains` | `public` | `Boolean` | `Int32 row, Int32 col` |
| `Contains` | `public` | `Boolean` | `CellRange range` |
| `GetCaptions` | `internal` | `IEnumerable`1` | `` |
| `GetCaptionText` | `public` | `String` | `` |
| `GetInnerCaption` | `public` | `String` | `` |
| `HasCaption` | `public` | `Boolean` | `Direction direction` |
| `SetAreaPer` | `public` | `Void` | `Double value` |
| `SetAspect` | `public` | `Void` | `Double value` |
| `SetCaption` | `public` | `Caption` | `Sheet sheet, Cell cell, List`1 captions` |
| `SetCaption` | `public` | `Void` | `String cellText` |

#### `Cimtops.R2Cluster.Cluster+<>c`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<GetCaptionText>b__51_0` | `internal` | `String` | `Caption c` |
| `<GetInnerCaption>b__52_0` | `internal` | `Boolean` | `Caption c` |

#### `Cimtops.R2Cluster.Cluster+<>c__DisplayClass46_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<SetCaption>b__0` | `internal` | `Boolean` | `Caption c` |

#### `Cimtops.R2Cluster.Cluster+<>c__DisplayClass53_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<HasCaption>b__0` | `internal` | `Boolean` | `Caption c` |

#### `Cimtops.R2Cluster.Cluster+<>c__DisplayClass54_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.R2Cluster.Cluster+<>c__DisplayClass54_1`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<AdustCaption>b__0` | `internal` | `Boolean` | `Cluster c` |

#### `Cimtops.R2Cluster.Cluster+<>c__DisplayClass54_2`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<AdustCaption>b__1` | `internal` | `Boolean` | `Cluster c` |

#### `Cimtops.R2Cluster.Decoder`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<IsCheckCluster>g__Func|35_0` | `private` | `Boolean` | `<>c__DisplayClass35_0& ` |
| `<SetInputParameter>g__Set|25_0` | `internal static` | `Void` | `String key, <>c__DisplayClass25_0& ` |
| `ArgbToText` | `private` | `String` | `Int32 argb` |
| `Cap2Kind` | `private static` | `Tuple`2` | `String caption` |
| `ClearTitle` | `public static` | `Void` | `String filePath, String sheetNae` |
| `CreateCluster` | `private` | `Cluster` | `Cell cell, Sheet sheet, IReadOnlyList`1 list, List`1 captions, ClusterInfo info, TitleInfo title` |
| `CreateColorDic` | `public static` | `Void` | `Tuple`2[] clusterColors, Tuple`2[] notClusterColors, Tuple`3[] clusterCaptions` |
| `DoAutomaticJudgement` | `public static` | `AutoJudgeResult` | `Book book, CancellationToken token, Boolean useAI, DeviceType device` |
| `DoAutomaticJudgement` | `private` | `List`1` | `Sheet sheet, Boolean useAI, AutoJudgeResult result, CancellationToken token, DeviceType device` |
| `FindTitle` | `public static` | `TitleInfo` | `String filePath, String sheetName` |
| `GetInstance` | `internal static` | `Decoder` | `String filePath, String sheetName` |
| `GetInstance` | `private static` | `Decoder` | `Sheet sheet` |
| `GetLikelihood` | `public static` | `Double` | `String path, String sheetName, IEnumerable`1 points, String typeKey` |
| `GetSheetTitle` | `private` | `TitleInfo` | `Sheet sheet` |
| `GetSheetTitleS` | `public static` | `TitleInfo` | `Sheet sheet` |
| `GetTitle` | `private` | `TitleInfo` | `Sheet sheet, Boolean& isCreatedNow` |
| `IsCheckCluster` | `private` | `Boolean` | `BorderQ borders, Cell cell, List`1 clusters` |
| `IsCheckSplitter` | `private` | `Boolean` | `String text` |
| `IsCluster` | `private` | `Boolean` | `Cell cell, List`1 clusters, Double allArea, DeviceType device, ClusterInfo& info, TitleInfo title, Dictionary`2 modify` |
| `IsNoConfidence` | `public static` | `AIInfo` | `String path, String sheetName, Point point` |
| `IsTableHead` | `private` | `Boolean` | `Cell cell, BorderQ borders` |
| `SetInputParameter` | `private` | `String` | `String src, String actionKind` |
| `SplitLine` | `private` | `IEnumerable`1` | `String text` |

#### `Cimtops.R2Cluster.Decoder+<>c`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<IsCluster>b__34_0` | `internal` | `Boolean` | `String reason` |
| `<IsCluster>b__34_1` | `internal` | `Boolean` | `String reason` |

#### `Cimtops.R2Cluster.Decoder+<>c__DisplayClass24_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<DoAutomaticJudgement>b__0` | `internal` | `Void` | `Int32 j, String text` |

#### `Cimtops.R2Cluster.Decoder+<>c__DisplayClass25_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.R2Cluster.Decoder+<>c__DisplayClass33_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<Cap2Kind>b__0` | `internal` | `Boolean` | `Tuple`3 c` |

#### `Cimtops.R2Cluster.Decoder+<>c__DisplayClass35_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.R2Cluster.Decoder+<>c__DisplayClass36_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<CreateCluster>g__SetAround|0` | `internal` | `Cell` | `Int32 row, Int32 col` |
| `<CreateCluster>g__SetSide|2` | `internal` | `Void` | `Int32 col` |
| `<CreateCluster>g__SetUpDown|1` | `internal` | `Void` | `Int32 row` |

#### `Cimtops.R2Cluster.Decoder+<>c__DisplayClass36_1`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<CreateCluster>b__3` | `internal` | `Boolean` | `Cluster c` |

#### `Cimtops.R2Cluster.Decoder+<>c__DisplayClass36_2`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<CreateCluster>b__4` | `internal` | `Boolean` | `Cluster c` |

#### `Cimtops.R2Cluster.Decoder+<SplitLine>d__29`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `MoveNext` | `private` | `Boolean` | `` |
| `System.Collections.Generic.IEnumerable<System.String>.GetEnumerator` | `private` | `IEnumerator`1` | `` |
| `System.Collections.IEnumerable.GetEnumerator` | `private` | `IEnumerator` | `` |
| `System.Collections.IEnumerator.Reset` | `private` | `Void` | `` |
| `System.IDisposable.Dispose` | `private` | `Void` | `` |

#### `Cimtops.R2Cluster.Decoder+AIInfo`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.R2Cluster.Decoder+AutoJudgeResult`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.R2Cluster.Decoder+ClusterInfo`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `CalcArea` | `public` | `Void` | `Cell cell` |

#### `Cimtops.R2Cluster.Global`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.R2Cluster.JudgeResult`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `Check` | `internal` | `Boolean` | `HttpClient hc, String sheetTitle, Cluster cl, Boolean useAI` |
| `Clear` | `public` | `Void` | `` |
| `Get` | `public` | `Double` | `String type` |
| `GetClusterKind` | `public static` | `String` | `String inner` |
| `GetTopKey` | `public` | `String` | `` |
| `IsNoConfidence` | `public` | `Boolean` | `` |
| `Set` | `internal` | `Void` | `String type, Double likelihood` |
| `Set` | `private` | `Void` | `Dictionary`2 dic, String key, String val` |
| `Setup` | `internal` | `Void` | `` |
| `UsedAI` | `public` | `Boolean` | `` |

#### `Cimtops.R2Cluster.JudgeResult+<>c__DisplayClass25_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<Check>b__0` | `internal` | `Boolean` | `Judement j` |

#### `Cimtops.R2Cluster.JudgeResult+<>c__DisplayClass4_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<GetClusterKind>b__0` | `internal` | `Boolean` | `Judement j` |

#### `Cimtops.R2Cluster.JudgeResult+Judement`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `CheckCaption` | `private` | `Boolean` | `String[] array, String cap` |
| `CheckMinMax` | `private` | `Boolean` | `Nullable`1 min, Nullable`1 max, Nullable`1 value` |
| `Get` | `private static` | `String` | `String[] csv, UInt32 index` |
| `GetSpan` | `private static` | `Nullable`1` | `String[] csv, UInt32 index, Boolean isLeft` |
| `IsCluster` | `public` | `Boolean` | `String inner` |
| `IsMatch` | `public` | `Boolean` | `Cluster cl, String cap, String inner` |
| `IsMatch` | `private` | `Boolean` | `String caption, String pattern` |

#### `Cimtops.R2Cluster.JudgeResult+Judement+<>c__DisplayClass32_0`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `<CheckCaption>b__0` | `internal` | `Boolean` | `String p` |

#### `Cimtops.R2Cluster.TitleInfo`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|

#### `Cimtops.R2Cluster.Util`

| Method | Access | Return | Parameters |
|--------|--------|--------|------------|
| `GetOrNew` | `public static` | `TValue` | `Dictionary`2 dic, TKey key` |
| `RemoveCaptionSpace` | `public static` | `Void` | `String& str` |

## iReporterExcelAddIn — Main AddIn

**File:** `C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn\iReporterExcelAddIn.dll`
**Size:** 74,240 bytes

**Architecture:** x86 (32-bit)
**Managed (.NET) assembly:** YES
**CLR Header:** RVA=0x00002008, Size=0x00000048

### Types

| Type | Kind | Public Methods | Public Fields |
|------|------|---------------|--------------|
**Reflection error:** ReflectionTypeLoadException: Unable to load one or more of the requested types.
Could not load file or assembly 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'. The system cannot find the file specified.
Could not load file or assembly 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'. The system cannot find the file specified.
Could not load file or assembly 'System.Configuration.ConfigurationManager, Version=0.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51'. The system cannot find the file specified.
Could not load file or assembly 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'. The system cannot find the file specified.

### String-Based Analysis (Fallback)

```
 2016 CIMTOPS CORPORATION
 c 2016 CIMTOPS CORPORATION
 CIMTOPS CORPORATION
 CORPORATION
 Version
_ERROR
_EXTRACTION
_INFO
_KIND
_MANY_SELECT
_NAME
_NAME2
_PRINT
_SELECT
_SETTING
_SHEET
_SHEET_NAME
_SHEET_NAME2
_TABLE
_VERSION_INFO
:3C4G5N
!"$#%&''(())6*82:3C4G5N
.AutoScaleBaseSize
.ClientSize
.Excel.Workbook
.Font
.Image
.Label
.Location
.Name
.OfficeMenu.Name
.OfficeMenu.Parent
.OfficeMenu.Type
.OfficeMenu.ZOrder
.Parent
.Properties.Resources
.Resources
.Size
.StartPosition
.TabIndex
.Text
.TextAlign
.Type
.Workbook
.ZOrder
''(())6*82:3C4G5N
'(())6*82:3C4G5N
"$#%&''(())6*82:3C4G5N
(())6*82:3C4G5N
())6*82:3C4G5N
))6*82:3C4G5N
)6*82:3C4G5N
*82:3C4G5N
&''(())6*82:3C4G5N
#%&''(())6*82:3C4G5N
%&''(())6*82:3C4G5N
>>$this.Name
>>$this.Type
>>autoJudgeButton.Name
>>autoJudgeButton.Parent
>>autoJudgeButton.Type
>>autoJudgeButton.ZOrder
>>button1.Name
>>button1.Parent
>>button1.Type
>>button1.ZOrder
>>checkSheetNameButton.Name
>>checkSheetNameButton.Parent
>>checkSheetNameButton.Type
>>checkSheetNameButton.ZOrder
>>clusterSettingButton.Name
>>clusterSettingButton.Parent
>>clusterSettingButton.Type
>>clusterSettingButton.ZOrder
>>colorSettingButton.Name
>>colorSettingButton.Parent
>>colorSettingButton.Type
>>colorSettingButton.ZOrder
>>group1.Name
>>group1.Parent
>>group1.Type
>>group1.ZOrder
>>labelMainMsg.Name
>>labelMainMsg.Parent
>>labelMainMsg.Type
>>labelMainMsg.ZOrder
>>Ribbon1.OfficeMenu.Name
>>Ribbon1.OfficeMenu.Parent
>>Ribbon1.OfficeMenu.Type
>>Ribbon1.OfficeMenu.ZOrder
>>tab1.Name
>>tab1.Type
>>tableSettingButton.Name
>>tableSettingButton.Parent
>>tableSettingButton.Type
>>tableSettingButton.ZOrder
>$this.Name
>$this.Type
>autoJudgeButton.Name
>autoJudgeButton.Parent
>autoJudgeButton.Type
>autoJudgeButton.ZOrder
>button1.Name
>button1.Parent
>button1.Type
>button1.ZOrder
>checkSheetNameButton.Name
>checkSheetNameButton.Parent
>checkSheetNameButton.Type
>checkSheetNameButton.ZOrder
>clusterSettingButton.Name
>clusterSettingButton.Parent
>clusterSettingButton.Type
>clusterSettingButton.ZOrder
>colorSettingButton.Name
>colorSettingButton.Parent
>colorSettingButton.Type
>colorSettingButton.ZOrder
>group1.Name
>group1.Parent
>group1.Type
>group1.ZOrder
>labelMainMsg.Name
>labelMainMsg.Parent
>labelMainMsg.Type
>labelMainMsg.ZOrder
>Ribbon1.OfficeMenu.Name
>Ribbon1.OfficeMenu.Parent
>Ribbon1.OfficeMenu.Type
>Ribbon1.OfficeMenu.ZOrder
>tab1.Name
>tab1.Type
>tableSettingButton.Name
>tableSettingButton.Parent
>tableSettingButton.Type
>tableSettingButton.ZOrder
$#%&''(())6*82:3C4G5N
$this.AutoScaleBaseSize
$this.ClientSize
$this.Name
$this.StartPosition
$this.Text
$this.Type
016 CIMTOPS CORPORATION
019.Properties.Resources
1.Label
1.Location
1.Name
1.OfficeMenu.Name
1.OfficeMenu.Parent
1.OfficeMenu.Type
1.OfficeMenu.ZOrder
1.Parent
1.Size
1.TabIndex
1.Text
1.Type
1.ZOrder
16 CIMTOPS CORPORATION
19.Properties.Resources
2:3C4G5N
2016 CIMTOPS CORPORATION
2019.Properties.Resources
3C4G5N
4G5N
6 CIMTOPS CORPORATION
6*82:3C4G5N
82:3C4G5N
9.Properties.Resources
ab1.Label
ab1.Name
ab1.Type
abelMainMsg
abelMainMsg.Font
abelMainMsg.Location
abelMainMsg.Name
abelMainMsg.Parent
abelMainMsg.Size
abelMainMsg.TabIndex
abelMainMsg.Text
abelMainMsg.TextAlign
abelMainMsg.Type
abelMainMsg.ZOrder
abIndex
ableSettingButton
ableSettingButton.Image
ableSettingButton.Label
ableSettingButton.Name
ableSettingButton.Parent
ableSettingButton.Type
ableSettingButton.ZOrder
Activate
AddIn
AddIn.dll
AddIn2019.Properties.Resources
ageDesc
ainMsg
ainMsg.Font
ainMsg.Location
ainMsg.Name
... and 1228 more
```
## iReporterExcelAddIn — Common

**File:** `C:\Program Files (x86)\CIMTOPS CORPORATION\iReporterExcelAddIn\iReporterExcelAddInCommon.dll`
**Size:** 375,808 bytes

**Architecture:** x86 (32-bit)
**Managed (.NET) assembly:** YES
**CLR Header:** RVA=0x00002008, Size=0x00000048

### Types

| Type | Kind | Public Methods | Public Fields |
|------|------|---------------|--------------|
**Reflection error:** ReflectionTypeLoadException: Unable to load one or more of the requested types.
Could not load file or assembly 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'. The system cannot find the file specified.
Could not load file or assembly 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'. The system cannot find the file specified.
Could not load file or assembly 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'. The system cannot find the file specified.
Could not load file or assembly 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load type 'System.Windows.DependencyObject' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'. The system cannot find the file specified.
Could not load file or assembly 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'. The system cannot find the file specified.
Could not load type 'System.Windows.DependencyObject' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load type 'System.Windows.DependencyObject' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.

### String-Based Analysis (Fallback)

```
 2015 CIMTOPS CORPORATION
 c 2015 CIMTOPS CORPORATION
 CIMTOPS CORPORATION
 CORPORATION
 Files(*.xml)|*.xml
 Gothic
 UI Gothic
 Version
_ACTION
_AUTO_INPUT_CLUSTER
_BATCHCLEAR
_BIOMETRICS
_BIOMETRICS+
_BTL
_BTLE
_CALC_FORMATS
_CAPTION
_CIOTG
_CIOTGZ
_CLUSTER
_CLUSTER_COL
_CLUSTERS
_CODE
_COL
_COL_KEY_EMPTY
_COL_KEY_INVALID
_COL_KEY_OVERLAP
_COL_NAME_EMPTY
_COL_OVERLAP
_COL_TYPE
_COMMAND
_COPY
_DATE
_DOCUMENT
_DOCUMENTk
_EMPTY
_ExcelAddIn
_FILE_LOAD
_FILE_SAVE
_FORMATS
_FOUND_CLUSTER
_FOUND_CLUSTERS
_FOUND_SHEET
_HAS_TABLE
_i-Reporter_ExcelAddIn
_INFO
_INPUT_CLUSTER
_INTERVAL
_INVALID
_JUMP
_KEY_EMPTY
_KEY_INVALID
_KEY_OVERLAP
_KIND
_KIND_AUTO_INPUT_CLUSTER
_KIND_BATCHCLEAR
_KIND_BIOMETRICS
_KIND_BIOMETRICS+
_KIND_BTL
_KIND_BTLE
_KIND_CIOTG
_KIND_CIOTGZ
_KIND_DOCUMENT
_KIND_DOCUMENTk
_KIND_MENU
_KIND_NO_NTFO
_KIND_OPEN_URL
_KIND_OUTPUT_TEXT
_KIND_QR_CODE
_KIND_RUN_COMMAND
_KIND_SHEET_COPY
_KIND_SHEET_JUMP
_KIND_START_TIMER
_LOAD
_MENU
_NAME_DATE
_NAME_EMPTY
_NAME_INTERVAL
_NAME_INVALID
_NAME_NUMERIC
_NAME_TEXT
_NO_CLUSTERS
_NO_INVALID
_NO_NTFO
_NO_OVERLAP
_NTFO
_NUM
_NUMERIC
_OPEN_URL
_OUTPUT_TEXT
_OUTPUT_TYPE
_OVERLAP
_QR_CODE
_ROW_OVERLAP
_RUN_COMMAND
_SAME_CLUSTER_COL
_SAVE
_SAVED
_SHEET
_SHEET_COPY
_SHEET_JUMP
_START_TIMER
_TABLE
_TEXT
_TIME
_TIMER
_TYPE
_TYPE_ACTION
_TYPE_DATE
_TYPE_NUM
_TYPE_TEXT
_TYPE_TIME
_URL
_VERSION_INFO
_zh-CN.xml
_zh-TW.xml
-?Pt}
-CN.xml
-Hans
-Hant
-Reporter_ExcelAddIn
-TW.xml
?Pt}
.Anchor
.ApplicationConfig.xml
.AutoScaleDimensions
.AutoScroll
.AutoSize
.ClientSize
.ColumnHeadersHeight
.CommonResource
.Domain.Resources.SheetNameInvalidCharacters
.Domain.Resources.SheetNameInvalidCharacters_ja.xml
.Domain.Resources.SheetNameInvalidCharacters_zh-CN.xml
.Domain.Resources.SheetNameInvalidCharacters_zh-TW.xml
.Domain.Resources.SheetNameInvalidCharacters.xml
.Domain.ValueDefinitions.Resources.CommonResource
.Enabled
.Font
.HeaderText
.ImeMode
.ItemHeight
.Location
.Margin
.Multiline
.Name
.Padding
.Parent
.Properties.Resources
.Resources
.Resources.CommonResource
.Resources.SheetNameInvalidCharacters
.Resources.SheetNameInvalidCharacters_ja.xml
.Resources.SheetNameInvalidCharacters_zh-CN.xml
.Resources.SheetNameInvalidCharacters_zh-TW.xml
.Resources.SheetNameInvalidCharacters.xml
.resources.ViewTexts
.ScrollBars
.SheetNameInvalidCharacters
.SheetNameInvalidCharacters_ja.xml
.SheetNameInvalidCharacters_zh-CN.xml
.SheetNameInvalidCharacters_zh-TW.xml
.SheetNameInvalidCharacters.xml
.Size
.StartPosition
.TabIndex
.Text
.TextAlign
.Type
.ValueDefinitions.Resources.CommonResource
.Views.resources.ViewTexts
.ViewTexts
.Visible
.Width
.ZOrder
(Simplified)
(Traditional)
(Version 
{F4}
/conmas/userConfig
/iReporterExcelAddInCommon;component/views/messagegrid.xaml
/userConfig
&-?Pt}
>>$this.Name
>>$this.Type
>>androidRadioButton.Name
>>androidRadioButton.Parent
>>androidRadioButton.Type
>>androidRadioButton.ZOrder
>>bothRadio.Name
>>bothRadio.Parent
>>bothRadio.Type
>>bothRadio.ZOrder
>>bottomHintCheckBox.Name
>>bottomHintCheckBox.Parent
>>bottomHintCheckBox.Type
>>bottomHintCheckBox.ZOrder
>>boxColKey.Name
>>boxColKey.Parent
>>boxColKey.Type
... and 11452 more
```
## ConMas Generator — GeneratorLib

**File:** `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator\ConMasGeneratorLib.dll`
**Size:** 114,176 bytes

**Architecture:** x86 (32-bit)
**Managed (.NET) assembly:** YES
**CLR Header:** RVA=0x00002008, Size=0x00000048

### Types

| Type | Kind | Public Methods | Public Fields |
|------|------|---------------|--------------|
| `ConMasGeneratorLib.Controller.BaseController` | Class | 2 | 0 |
| `ConMasGeneratorLib.Controller.ProcessController` | Class | 1 | 0 |
| `ConMasGeneratorLib.Controller.ProcessController+<>c` | Class | 0 | 4 |
| `ConMasGeneratorLib.Controller.Rest` | Class | 2 | 1 |
| `ConMasGeneratorLib.Controller.RestController` | Class | 2 | 1 |
| `ConMasGeneratorLib.Controller.WebClientEx` | Class | 0 | 0 |
| `ConMasGeneratorLib.Data.AppSettingData` | Class | 0 | 0 |
| `ConMasGeneratorLib.Data.Command` | Class | 0 | 0 |
| `ConMasGeneratorLib.Data.GeneratorData` | Class | 1 | 0 |
| `ConMasGeneratorLib.Data.JobData` | Class | 1 | 2 |
| `ConMasGeneratorLib.Data.ParameterData` | Class | 0 | 0 |
| `ConMasGeneratorLib.Data.ProcessData` | Class | 1 | 2 |
| `ConMasGeneratorLib.Data.ProcessResultData` | Class | 3 | 11 |
| `ConMasGeneratorLib.DB.JobDbLog` | Class | 3 | 0 |
| `ConMasGeneratorLib.DB.LocalDB` | Class | 2 | 0 |
| `ConMasGeneratorLib.GeneratorMain` | Class | 2 | 0 |
| `ConMasGeneratorLib.Global` | Class | 0 | 1 |
| `ConMasGeneratorLib.Job.JobBase` | Class | 3 | 1 |
| `ConMasGeneratorLib.Job.JobBase+<>c` | Class | 0 | 2 |
| `ConMasGeneratorLib.Job.JobController` | Class | 11 | 0 |
| `ConMasGeneratorLib.Job.JobController+<>c` | Class | 0 | 3 |
| `ConMasGeneratorLib.Job.JobController+<>c__DisplayClass23_0` | Class | 0 | 2 |
| `ConMasGeneratorLib.Job.JobController+<>c__DisplayClass25_0` | Class | 0 | 2 |
| `ConMasGeneratorLib.Job.JobController+JobType` | Enum | 0 | 3 |
| `ConMasGeneratorLib.Job.Processes.Process` | Class | 1 | 4 |
| `ConMasGeneratorLib.Job.Schdule.JobSchdule` | Class | 2 | 0 |
| `ConMasGeneratorLib.Job.Watcher.JobWatcher` | Class | 3 | 0 |
| `ConMasGeneratorLib.Properties.Resources` | Class | 0 | 0 |
| `ConMasGeneratorLib.TaskScheduler.DailyTrigger` | Struct | 0 | 0 |
| `ConMasGeneratorLib.TaskScheduler.DaysOfWeek` | Enum | 0 | 9 |
| `ConMasGeneratorLib.TaskScheduler.IdleWait` | Struct | 0 | 0 |
| `ConMasGeneratorLib.TaskScheduler.Interop.IEnumWorkItems` | Interface | 4 | 0 |
| `ConMasGeneratorLib.TaskScheduler.Interop.ITask` | Interface | 41 | 0 |
| `ConMasGeneratorLib.TaskScheduler.Interop.ITaskScheduler` | Interface | 8 | 0 |
| `ConMasGeneratorLib.TaskScheduler.Interop.ITaskTrigger` | Interface | 3 | 0 |
| `ConMasGeneratorLib.TaskScheduler.Interop.TaskClass` | Class | 0 | 0 |
| `ConMasGeneratorLib.TaskScheduler.Interop.TaskSchedulerClass` | Class | 0 | 0 |
| `ConMasGeneratorLib.TaskScheduler.Interop.TriggerTypeUnion` | Struct | 0 | 0 |
| `ConMasGeneratorLib.TaskScheduler.MonthlyDateTrigger` | Struct | 4 | 0 |
| `ConMasGeneratorLib.TaskScheduler.MonthlyDayOfWeekTrigger` | Struct | 0 | 0 |
| `ConMasGeneratorLib.TaskScheduler.Months` | Enum | 0 | 14 |
| `ConMasGeneratorLib.TaskScheduler.Task` | Class | 18 | 1 |
| `ConMasGeneratorLib.TaskScheduler.TaskFlags` | Enum | 0 | 15 |
| `ConMasGeneratorLib.TaskScheduler.TaskSchedule` | Class | 4 | 0 |
| `ConMasGeneratorLib.TaskScheduler.TaskStatus` | Enum | 0 | 4 |
| `ConMasGeneratorLib.TaskScheduler.TaskTrigger` | Class | 13 | 0 |
| `ConMasGeneratorLib.TaskScheduler.TriggerFlags` | Enum | 0 | 5 |
| `ConMasGeneratorLib.TaskScheduler.TriggerType` | Enum | 0 | 9 |
| `ConMasGeneratorLib.TaskScheduler.Unmanaged.ComHelpers.ComInterfaceWrapper` | Class | 2 | 0 |
| `ConMasGeneratorLib.TaskScheduler.Unmanaged.ComHelpers.HResult` | Enum | 0 | 21 |
| `ConMasGeneratorLib.TaskScheduler.Unmanaged.CoTaskMem` | Class | 3 | 0 |
| `ConMasGeneratorLib.TaskScheduler.Unmanaged.SimpleMemory` | Abstract Class | 6 | 0 |
| `ConMasGeneratorLib.TaskScheduler.Unmanaged.SystemTimeClass` | Class | 6 | 2 |
| `ConMasGeneratorLib.TaskScheduler.Unmanaged.UnmanagedMemory` | Abstract Class | 9 | 0 |
| `ConMasGeneratorLib.TaskScheduler.WeeklyTrigger` | Struct | 0 | 0 |
| `ConMasGeneratorLib.TaskScheduler.WhichWeek` | Enum | 0 | 6 |
| `ConMasGeneratorLib.UserException.BaseException` | Class | 0 | 0 |
| `ConMasGeneratorLib.UserException.ErrorConst` | Class | 0 | 2 |
| `ConMasGeneratorLib.UserException.JobControllerException` | Class | 0 | 0 |
| `ConMasGeneratorLib.UserException.ProcException` | Class | 0 | 0 |
| `ConMasGeneratorLib.UserException.SchException` | Class | 0 | 0 |
| `ConMasGeneratorLib.UserException.WatcherException` | Class | 0 | 0 |
| `ConMasGeneratorLib.Util.Mail` | Class | 3 | 0 |
| `ConMasGeneratorLib.Util.MailSendData` | Class | 0 | 0 |
| `ConMasGeneratorLib.Util.Utility` | Class | 38 | 11 |
| `TaskScheduler._TASK_ACTION_TYPE` | Enum | 0 | 5 |
| `TaskScheduler._TASK_LOGON_TYPE` | Enum | 0 | 8 |
| `TaskScheduler._TASK_RUNLEVEL` | Enum | 0 | 3 |
| `TaskScheduler._TASK_TRIGGER_TYPE2` | Enum | 0 | 13 |
| `TaskScheduler.IAction` | Interface | 0 | 0 |
| `TaskScheduler.IActionCollection` | Interface | 1 | 0 |
| `TaskScheduler.IDailyTrigger` | Interface | 0 | 0 |
| `TaskScheduler.IExecAction` | Interface | 0 | 0 |
| `TaskScheduler.IIdleSettings` | Interface | 0 | 0 |
| `TaskScheduler.IMonthlyTrigger` | Interface | 0 | 0 |
| `TaskScheduler.IPrincipal` | Interface | 0 | 0 |
| `TaskScheduler.IRegisteredTask` | Interface | 0 | 0 |
| `TaskScheduler.IRegisteredTaskCollection` | Interface | 1 | 0 |
| `TaskScheduler.IRegistrationInfo` | Interface | 0 | 0 |
| `TaskScheduler.IRepetitionPattern` | Interface | 0 | 0 |
| `TaskScheduler.ITaskDefinition` | Interface | 0 | 0 |
| `TaskScheduler.ITaskFolder` | Interface | 3 | 0 |
| `TaskScheduler.ITaskService` | Interface | 3 | 0 |
| `TaskScheduler.ITaskSettings` | Interface | 0 | 0 |
| `TaskScheduler.ITimeTrigger` | Interface | 0 | 0 |
| `TaskScheduler.ITrigger` | Interface | 0 | 0 |
| `TaskScheduler.ITriggerCollection` | Interface | 1 | 0 |
| `TaskScheduler.IWeeklyTrigger` | Interface | 0 | 0 |
| `TaskScheduler.TaskScheduler` | Interface | 0 | 0 |

## ConMas Generator — Job Executable

**File:** `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator\ConMasJob.exe`
**Size:** 8,704 bytes

**Architecture:** x86 (32-bit)
**Managed (.NET) assembly:** YES
**CLR Header:** RVA=0x00002008, Size=0x00000048

**Reflection:** Could not load assembly. Falling back to string analysis.

### String-Based Analysis (Fallback)

```
 (c) CIMTOPS CORPORATION 2012
 CIMTOPS CORPORATION 2012
 CORPORATION
 CORPORATION 2012
 ERROR
 File Nothing ERROR
 JOBID ERROR
 JOBTYPE ERROR
 Nothing ERROR
 Version
_INFO
_VERSION_INFO
.Properties.Resources
.Resources
(c) CIMTOPS CORPORATION 2012
) CIMTOPS CORPORATION 2012
alCopyright
alFilename
alName
alTrademarks
anyName
arFileInfo
asJob
asJob.exe
asJob.Properties.Resources
Assembly Version
atch File Nothing ERROR
ATION 2012
b.Properties.Resources
BID ERROR
bly Version
BTYPE ERROR
c) CIMTOPS CORPORATION 2012
ch File Nothing ERROR
CIMTOPS CORPORATION
CIMTOPS CORPORATION 2012
Comments
CompanyName
ConMasJob
ConMasJob.exe
ConMasJob.Properties.Resources
Copyright
Copyright (c) CIMTOPS CORPORATION 2012
CORPORATION 2012
ctName
ctVersion
D ERROR
Description
ductName
ductVersion
E ERROR
e Nothing ERROR
eDescription
egalCopyright
egalTrademarks
eInfo
embly Version
ernalName
ERSION_INFO
erties.Resources
es.Resources
eVersion
File Nothing ERROR
FileDescription
FileInfo
Filename
FileVersion
g ERROR
g JOBID ERROR
g JOBTYPE ERROR
galCopyright
galTrademarks
gFileInfo
ght (c) CIMTOPS CORPORATION 2012
ginalFilename
h File Nothing ERROR
hing ERROR
hing JOBID ERROR
hing JOBTYPE ERROR
ht (c) CIMTOPS CORPORATION 2012
ID ERROR
ies.Resources
ight (c) CIMTOPS CORPORATION 2012
iginalFilename
ile Nothing ERROR
ileDescription
ileInfo
ileVersion
IMTOPS CORPORATION
IMTOPS CORPORATION 2012
inalFilename
Info
ing ERROR
ing JOBID ERROR
ing JOBTYPE ERROR
ingFileInfo
InternalName
ION 2012
ION_INFO
Job.exe
Job.Properties.Resources
JOBID ERROR
JOBTYPE ERROR
lCopyright
le Nothing ERROR
leDescription
LegalCopyright
LegalTrademarks
leInfo
leVersion
lFilename
lName
lTrademarks
ly Version
MasJob
MasJob.exe
MasJob.Properties.Resources
mbly Version
mpanyName
MTOPS CORPORATION
MTOPS CORPORATION 2012
N 2012
N_INFO
nalFilename
nalName
Name
ng ERROR
ng JOBID ERROR
ng JOBTYPE ERROR
ngFileInfo
nMasJob
nMasJob.exe
nMasJob.Properties.Resources
Nothing ERROR
Nothing JOBID ERROR
Nothing JOBTYPE ERROR
nternalName
nyName
ob.Properties.Resources
OBID ERROR
OBTYPE ERROR
oductName
oductVersion
ompanyName
ON 2012
ON_INFO
onMasJob
onMasJob.exe
onMasJob.Properties.Resources
operties.Resources
OPS CORPORATION
OPS CORPORATION 2012
opyright (c) CIMTOPS CORPORATION 2012
ORATION 2012
OriginalFilename
ORPORATION 2012
othing ERROR
othing JOBID ERROR
othing JOBTYPE ERROR
panyName
PE ERROR
perties.Resources
PORATION 2012
ProductName
ProductVersion
Properties.Resources
PS CORPORATION
PS CORPORATION 2012
pyright (c) CIMTOPS CORPORATION 2012
RATION 2012
Resources
rFileInfo
right (c) CIMTOPS CORPORATION 2012
riginalFilename
ringFileInfo
rnalName
roductName
roductVersion
roperties.Resources
RPORATION 2012
RSION_INFO
rties.Resources
S CORPORATION
S CORPORATION 2012
S_VERSION_INFO
s.Resources
sembly Version
SION_INFO
sJob
sJob.exe
sJob.Properties.Resources
ssembly Version
StringFileInfo
t (c) CIMTOPS CORPORATION 2012
tch File Nothing ERROR
ternalName
thing ERROR
thing JOBID ERROR
thing JOBTYPE ERROR
ties.Resources
... and 20 more
```
## ConMas Generator — Tool Executable

**File:** `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator\ConMasTool.exe`
**Size:** 7,168 bytes

**Architecture:** x86 (32-bit)
**Managed (.NET) assembly:** YES
**CLR Header:** RVA=0x00002008, Size=0x00000048

**Reflection:** Could not load assembly. Falling back to string analysis.

### String-Based Analysis (Fallback)

```
 Version
_INFO
_JIS
_VERSION_INFO
.Properties.Resources
.Resources
a-JP
alCopyright
alFilename
alName
alTrademarks
anyName
arFileInfo
Assembly Version
asTool
asTool.exe
asTool.Properties.Resources
bly Version
Comments
CompanyName
ConMasTool
ConMasTool.exe
ConMasTool.Properties.Resources
Copyright
Copyright 
ctName
ctVersion
Description
ductName
ductVersion
eDescription
egalCopyright
egalTrademarks
eInfo
embly Version
ernalName
ERSION_INFO
erties.Resources
es.Resources
eVersion
FileDescription
FileInfo
Filename
FileVersion
ft_JIS
galCopyright
galTrademarks
gFileInfo
ginalFilename
hift_JIS
ies.Resources
ift_JIS
iginalFilename
ileDescription
ileInfo
ileVersion
inalFilename
Info
ingFileInfo
InternalName
ION_INFO
ja-JP
l.Properties.Resources
lCopyright
leDescription
LegalCopyright
LegalTrademarks
leInfo
leVersion
lFilename
lName
lTrademarks
ly Version
MasTool
MasTool.exe
MasTool.Properties.Resources
mbly Version
mpanyName
N_INFO
nalFilename
nalName
Name
ngFileInfo
nMasTool
nMasTool.exe
nMasTool.Properties.Resources
nternalName
nyName
oductName
oductVersion
ol.Properties.Resources
ompanyName
ON_INFO
onMasTool
onMasTool.exe
onMasTool.Properties.Resources
ool.Properties.Resources
operties.Resources
OriginalFilename
panyName
perties.Resources
ProductName
ProductVersion
Properties.Resources
Resources
rFileInfo
riginalFilename
ringFileInfo
rnalName
roductName
roductVersion
roperties.Resources
RSION_INFO
rties.Resources
S_VERSION_INFO
s.Resources
sembly Version
Shift_JIS
SION_INFO
ssembly Version
sTool
sTool.exe
sTool.Properties.Resources
StringFileInfo
t_JIS
ternalName
ties.Resources
tName
Tool
Tool.exe
Tool.Properties.Resources
Trademarks
Translation
tringFileInfo
tVersion
uctName
uctVersion
VarFileInfo
Version
VERSION_INFO
VS_VERSION_INFO
y Version
yName
```
## ConMas Designer — Main Client

**File:** `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\bin\ConMasClient.exe`
**Size:** 3,882,544 bytes

**Architecture:** x86 (32-bit)
**Managed (.NET) assembly:** YES
**CLR Header:** RVA=0x00002008, Size=0x00000048

### Types

| Type | Kind | Public Methods | Public Fields |
|------|------|---------------|--------------|
**Reflection error:** ReflectionTypeLoadException: Unable to load one or more of the requested types.
Could not load type 'System.Windows.Markup.InternalTypeHelper' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'System.Configuration.ConfigurationManager, Version=0.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load type 'System.Windows.Freezable' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load file or assembly 'PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'. The system cannot find the file specified.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.

### String-Based Analysis (Fallback)

```
  Index[
 - [clusterId]
 - [displayValue]
 - [editTime]
 - [editUser]
 - [editUserName]
 - [parameterType]
 - [sheetNo]
 "os":"","appVersion":"","terminalId":"","userId":"","userName":"","clusters":[{"sheetNo":"1","clusterId":"0","name":"Cluster1","type":"Select","value":"","displayValue":"","editUser":"","editUserName":"","editTime":""},{"sheetNo":"2","clusterId":"3","name":"Cluster3","type":"Image","value":"","displayValue":"","editUser":"","editUserName":"","editTime":""}]}
 ({0}) Before: {1}, After: {2}
 (c) CIMTOPS CORPORATION 2011
 [clusterId]
 [displayValue]
 [editTime]
 [editUser]
 [editUserName]
 [parameterType]
 [sheetNo]
 {0}, Index = {1} }}
 {0}, Value = {1} }}
 {1}, After: {2}
 = {0}, Index = {1} }}
 = {0}, Value = {1} }}
 After: {2}
 and the number of indexes in TopReport do not match.
 Before: {1}, After: {2}
 CIMTOPS CORPORATION 2011
 Cluster
 CORPORATION
 CORPORATION 2011
 Count
 created on Designer
 Designer
 Designer 
 Designer Login
 Designer/{0}/
 Error ({0}) ({1}) ({2} expected)
 Error.
 Error. header: 
 Excel data is empty.
 Excel file does not exist.
 Excel file is incorrect.
 Excel file: 
 File (*.pdf,*.xps,*.gif,*.tif,*.tiff)|*.pdf;*.xps;*.gif;*.tif;*.tiff
 File (*.xls,*.xlsx)|*.xls;*.xlsx
 Files(*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg
 format of the Excel file is incorrect.
 HH:mm:ss
 Implemented
 Import
 in TopReport do not match.
 Index = {1} }}
 Index[
 indexes in TopReport do not match.
 insertAtPageNumber is out of range.
 ListTableData = {0} }}
 Login
 MultiPageImage
 Name = {0}, Value = {1} }}
 number of indexes in TopReport do not match.
 number of sheets and the number of indexes in TopReport do not match.
 of indexes in TopReport do not match.
 of sheets and the number of indexes in TopReport do not match.
 of the Excel file is incorrect.
 on Designer
 open Excel file: 
 Page File (*.pdf,*.xps,*.gif,*.tif,*.tiff)|*.pdf;*.xps;*.gif;*.tif;*.tiff
 PDF file does not exist.
 PDF file stream does not exist.
 PDF Version Error.
 PDF Version Error. header: 
 Reference calculation begins
 Reference calculation ends
 replaceAtPageNumber is out of range.
 Rev.{2}
 Rev.{2},{3}
 Row = {0}, Index = {1} }}
 sheets and the number of indexes in TopReport do not match.
 specified Excel file does not exist.
 the Excel file is incorrect.
 the number of indexes in TopReport do not match.
 to open Excel file: 
 TopReport do not match.
 traditionalChinese
 TreeData = {0} }}
 Value = {1} }}
 Version
 Version Error.
 Version Error. header: 
__AutoAlignmentThreshold
__GridFont
__SelectedClusterColor
__StandardClusterColor
__TextViewFont
_AllYes
_Area
_AutoAlignmentThreshold
_BarcodeLengthError
_Cancel
_Caption
_Checkmark
_CircledLargestNumber
_CircledNumber
_Click
_ClusterType
_CreateDelegate
_DATECLUSTER
_DUMMY_FUNCTION
_DuplicateError
_DuplicateErrorForTheSameSession
_Evidence
_Export
_FOCUS
_FUNCTION
_GATEWAY
_GridFont
_Import
_INFO
_JIS
_LIMIT
_Loaded
_MatchingError
_name":"AND"
_name":"OR",
_name":"SQRT"
_No"
_NOFILL
_NotAbleToOutputError
_Ok-
_OPERATIONS
_PinImage
_REFERENCE_LIMIT
_RepeatingMode0_BarcodeLengthError
_RepeatingMode0_Checkmark
_RepeatingMode0_CircledLargestNumber
_RepeatingMode0_CircledNumber
_RepeatingMode0_DuplicateError
_RepeatingMode0_DuplicateErrorForTheSameSession
_RepeatingMode0_MatchingError
_RepeatingMode0_Scannable
_RepeatingMode1_BarcodeLengthError
_RepeatingMode1_MatchingError
_RepeatingMode1_NotAbleToOutputError
_RepeatingMode1_Scannable
_REPO_SCAN
_RunWorkerCompleted has some errors
_SCAN
_Scannable
_ScanSetting_Export
_ScanSetting_Import
_SelectedClusterColor
_StandardClusterColor
_SUPPORT_DATECLUSTER
_TextViewFont
_TimeCalculate
_VERSION_INFO
_WAV
_windowTitle
_xlnm.Print_Area
_Yes
_Yes1
- [clusterId]
- [displayValue]
- [editTime]
- [editUser]
- [editUserName]
- [parameterType]
- [sheetNo]
-0000-0000-C000-000000000046
-0000-C000-000000000046
-1</itemIndex>
-9]*[cC][0-9]*
-C000-000000000046
-MM-dd
-MM-yy
-MM-yyyy
-Reporter
-Za-z]
-Za-z]:\\.*
-Za-z]+)([0-9]+)
-zA-z]+[0-9]+
-Za-z0-9]
, After: {2}
, Index = {1} }}
, Index[
, Rev.{2}
, Rev.{2},{3}
, Value = {1} }}
,,,/ConMasClient;component/Images/hh.png
,,,/ConMasClient;component/Images/vv.png
,,/ConMasClient;component/Images/hh.png
,,/ConMasClient;component/Images/vv.png
,"appVersion":"","terminalId":"","userId":"","userName":"","clusters":[{"sheetNo":"1","clusterId":"0","name":"Cluster1","type":"Select","value":"","displayValue":"","editUser":"","editUserName":"","editTime":""},{"sheetNo":"2","clusterId":"3","name":"Cluster3","type":"Image","value":"","displayValue":"","editUser":"","editUserName":"","editTime":""}]}
,"clusterId":"0","name":"Cluster1","type":"Select","value":"","displayValue":"","editUser":"","editUserName":"","editTime":""},{"sheetNo":"2","clusterId":"3","name":"Cluster3","type":"Image","value":"","displayValue":"","editUser":"","editUserName":"","editTime":""}]}
,"clusterId":"3","name":"Cluster3","type":"Image","value":"","displayValue":"","editUser":"","editUserName":"","editTime":""}]}
,"clusters":[{"sheetNo":"1","clusterId":"0","name":"Cluster1","type":"Select","value":"","displayValue":"","editUser":"","editUserName":"","editTime":""},{"sheetNo":"2","clusterId":"3","name":"Cluster3","type":"Image","value":"","displayValue":"","editUser":"","editUserName":"","editTime":""}]}
,"displayValue":"","editUser":"","editUserName":"","editTime":""},{"sheetNo":"2","clusterId":"3","name":"Cluster3","type":"Image","value":"","displayValue":"","editUser":"","editUserName":"","editTime":""}]}
,"displayValue":"","editUser":"","editUserName":"","editTime":""}]}
,"editTime":""},{"sheetNo":"2","clusterId":"3","name":"Cluster3","type":"Image","value":"","displayValue":"","editUser":"","editUserName":"","editTime":""}]}
,"editTime":""}]}
... and 14378 more
```
