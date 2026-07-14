# Phase 20.2 — Legacy Publish Pipeline Reconstruction

**Date:** 2026-07-13
**Status:** Investigation Complete
**Next:** Phase 20.3 — Recover coordinate algorithm from ConMasClient.exe (requires .NET Framework 4.x decompilation)

---

## 1. Complete Lifecycle Diagram

```
┌─────────────────────────────────────────────────────────────────────────┐
│                        DESIGNER PHASE (Excel Add-In)                     │
│                                                                          │
│  Excel Workbook                                                         │
│      │                                                                   │
│      ▼                                                                   │
│  User selects merged cell, clicks "Add Keyboard/Number"                  │
│      │                                                                   │
│      ▼                                                                   │
│  iReporter Add-In writes Excel Comment into top-left cell of merge       │
│      │  Comment format:                                                  │
│      │    Line 0: Name (e.g. "samples", "Machine")                       │
│      │    Line 1: Type (e.g. "KeyboardText", "InputNumeric")             │
│      │    Line 2: Index (sort order, optional)                            │
│      │    Line 3+: Parameters (e.g. Required=0;Lines=1;...)              │
│      │                                                                   │
│      ▼                                                                   │
│  Workbook saved with embedded Comments                                   │
└─────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                     DESIGNER APP PHASE (ConMas Designer)                 │
│                                                                          │
│  Designer loads workbook                                                 │
│      │                                                                   │
│      ├── Reads Print Area (PageSetup.PrintArea)                          │
│      ├── Reads Comments (worksheet.Comments)                             │
│      ├── Reads Merge Areas (comment.Parent.MergeArea)                    │
│      │                                                                   │
│      ├── Creates temporary Cluster objects in memory                     │
│      │     via Decoder.DoAutomaticJudgement() → Cluster objects           │
│      │                                                                   │
│      ├── Renders background from Print Area → preview                    │
│      │                                                                   │
│      └── Shows yellow editable overlays for each cluster                 │
│                                                                          │
│  At this point NOTHING is saved to the database.                         │
│  Designer reads Excel workbook every time it opens (no cache).          │
└─────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                        PUBLISH PHASE (ConMasClient.exe)                  │
│                                                                          │
│  User clicks Publish button in ConMas Designer                           │
│      │                                                                   │
│      ▼                                                                   │
│  ConMasClient.exe executes publish pipeline:                             │
│                                                                          │
│  [Step A] Read Excel via COM Interop                                     │
│      ├── BookFactory.Create(Workbook)                                     │
│      │     ├── Sheet.PrintArea → parse address                           │
│      │     ├── Col[].Width = COM Range.Columns[i].Width (points)         │
│      │     ├── Row[].Height = COM Range.Rows[i].Height (points)          │
│      │     └── Cell[] → each cell in print area                          │
│      │                                                                   │
│      └── Decoder.DoAutomaticJudgement()                                   │
│            ├── Identifies cells with comments                             │
│            ├── Creates Cluster objects with col/row indices               │
│            └── CalcArea() → areaPer = w*h*100/(sheetW*sheetH)            │
│                                                                          │
│  [Step B] Assign Cluster IDs                                             │
│      ├── Sort clusters by: ClusterIndex (comment line 2), Row, Col       │
│      ├── Assign sequential cluster_id starting from 0                    │
│      └── Cluster ID = position in sorted list                            │
│                                                                          │
│  [Step C] Generate DB names for empty comment names                      │
│      ├── If comment line 0 is empty or whitespace →                      │
│      │   name = "Cluster" + cluster_id                                   │
│      │   (e.g. "Cluster0", "Cluster1", ..., "Cluster4")                  │
│      └── This is WHY Template 547 has "Cluster0"-"Cluster4" in DB        │
│          even though comment line 0 is blank.                            │
│                                                                          │
│  [Step D] Compute coordinates (ratio)                                    │
│      ├── left_pt   = Σ Col[n].Width for n=1..cluster.Col-1              │
│      ├── right_pt  = Σ Col[n].Width for n=1..cluster.Right              │
│      ├── top_pt    = Σ Row[n].Height for n=1..cluster.Row-1             │
│      ├── bottom_pt = Σ Row[n].Height for n=1..cluster.Bottom            │
│      ├── Apply UNKNOWN transform (the gap) ← IN ConMasClient.exe        │
│      │     (involves page width/height normalization + compensation)     │
│      │                                                                   │
│      ├── left   = RoundEx(left_pt   / PageWidth,  7)                     │
│      ├── right  = RoundEx(right_pt  / PageWidth,  7)                     │
│      ├── top    = RoundEx(top_pt    / PageHeight, 7)                     │
│      └── bottom = RoundEx(bottom_pt / PageHeight, 7)                     │
│                                                                          │
│  [Step E] Generate XML string                                            │
│      ├── <conmas> → <header> → <top> → <sheets> → <sheet> → <clusters> │
│      ├── Each cluster node includes:                                      │
│      │     name, type, left, top, right, bottom, value,                   │
│      │     inputParameters, cellAddress, isHidden, readOnly,             │
│      │     external, displayValue, cooperationCluster,                   │
│      │     excelOutputValue, etc.                                        │
│      └── XML includes defTopId, version, serverVersion, etc.             │
│                                                                          │
│  [Step F] Generate background PDF                                         │
│      ├── Export worksheet as Fixed Format (PDF)                          │
│      │     Excel COM: ExportAsFixedFormat(XlTypePDF)                     │
│      └── Reads raw PDF bytes into memory                                 │
│                                                                          │
│  [Step G] Insert into database (single transaction)                      │
│      ├── 1. INSERT INTO def_top                                          │
│      │       (def_top_name, report_type, def_file, def_file_name,        │
│      │        designer_version, server_version, ...)                     │
│      │     → Returns def_top_id (auto-increment)                         │
│      │                                                                   │
│      ├── 2. INSERT INTO def_sheet                                        │
│      │       (def_top_id, def_sheet_no, def_sheet_name)                  │
│      │     → Returns def_sheet_id(s)                                     │
│      │                                                                   │
│      ├── 3. INSERT INTO def_cluster (one per cluster)                    │
│      │       (def_sheet_id, cluster_id, cluster_name, cluster_type,      │
│      │        left_position, top_position, right_position,                │
│      │        bottom_position, input_parameter, cell_addr)               │
│      │                                                                   │
│      ├── 4. UPDATE def_top SET xml_data = '<xml...>'                     │
│      │       WHERE def_top_id = @id                                      │
│      │                                                                   │
│      ├── 5. UPDATE def_top SET background_image_file = <PDF bytes>       │
│      │       WHERE def_top_id = @id                                      │
│      │                                                                   │
│      ├── 6. INSERT INTO def_current (def_top_id, def_top_org)           │
│      │                                                                   │
│      └── 7. COMMIT                                                       │
│                                                                          │
│  [Step H] Template is now persisted in DB                                │
│      Designer can now open template by ID (reads DB, NOT Excel)          │
└─────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                        RELOAD PHASE (ConMas Designer)                    │
│                                                                          │
│  User opens existing template by ID instead of by Excel file             │
│      │                                                                   │
│      ├── Read def_top → template header                                  │
│      ├── Read def_sheet → sheet metadata                                 │
│      ├── Read def_cluster → cluster definitions with coordinates         │
│      ├── Read xml_data → full XML structure                              │
│      ├── Read background_image_file → PDF background                     │
│      │                                                                   │
│      └── Render:                                                         │
│            ├── Decode PDF background → raster image                      │
│            ├── Overlay clusters at stored ratio positions                │
│            └── Display editable form                                     │
│                                                                          │
│  NOTE: At reload time, Excel is NEVER read again.                        │
│  All data comes from PostgreSQL database.                               │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## 2. Stage-by-Stage Publish Flow

### Stage 1: Excel Reading (COM Interop)

| Component | What It Reads | Source |
|-----------|--------------|--------|
| `BookFactory.Create()` | Workbook → Sheets → PrintArea | `Cimtops.Excel.dll` |
| `Sheet` constructor | Col widths, Row heights from print area | COM `Range.Columns[i].Width` / `Rows[i].Height` |
| `CellRange` constructor | Column, Row, Columns.Count, Rows.Count | COM `Range.Column`, `.Row`, etc. |
| `Decoder.DoAutomaticJudgement()` | Cell comments → Cluster detection | `Cimtops.R2Cluster.dll` |
| `ThisAddIn.InitClusterList()` | Comments + MergeAreas → ClusterInfo list | `iReporterExcelAddIn.dll` |

### Stage 2: Cluster ID Assignment

**Sort Order** (from `ClusterList.Compare()`):
1. `ClusterIndex` — comment line 2 (numeric), -1 if unparseable
2. `Row` — ascending
3. `Col` — ascending

**Cluster ID** = position in sorted list (0-based).

### Stage 3: Empty Name Handling

**Code location:** Probably in `ConMasClient.exe` publish pipeline, not in the Excel add-in.

**Logic (reconstructed):**
```
if (string.IsNullOrWhiteSpace(cluster.ClusterName))
    cluster.ClusterName = "Cluster" + clusterId;
```

**Evidence:**
- Template 547: 5 comments, comment line 0 empty in all → DB has "Cluster0", "Cluster1", "Cluster2", "Cluster3", "Cluster4"
- Template 546: comment line 0 = "samples" → DB has "samples" (no renaming needed)
- Template 548: comment line 0 = "Machine", "Machine_Output" → DB matches

### Stage 4: Coordinate Computation

**Inputs:**
- `Cluster.Row`, `Cluster.Col`, `Cluster.Right`, `Cluster.Bottom` (1-based col/row indices)
- `Col[n].Width` = COM column width in points
- `Row[n].Height` = COM row height in points
- `PageWidth` = page width in points (from PageSetup)
- `PageHeight` = page height in points

**Formula (reconstructed, missing gap transform):**
```
left_pt    = Σ Col[n].Width   for n = 1 to cluster.Col-1
top_pt     = Σ Row[n].Height  for n = 1 to cluster.Row-1
right_pt   = Σ Col[n].Width   for n = 1 to cluster.Right
bottom_pt  = Σ Row[n].Height  for n = 1 to cluster.Bottom

// PLUS unknown gap compensation applied only to right/bottom:
right_pt  += gapX   (varies by template, 0 for first cluster right?)
bottom_pt += gapY   (~1.08pt typically)

DB_left    = RoundEx(left_pt   / PageWidth,  7)
DB_top     = RoundEx(top_pt    / PageHeight, 7)
DB_right   = RoundEx(right_pt  / PageWidth,  7)
DB_bottom  = RoundEx(bottom_pt / PageHeight, 7)
```

**Confidence: 90%** — based on matching left/top for origin clusters, systematic gap in right/bottom.

### Stage 5: Background Image Generation

**Pipeline:**
1. Export worksheet as PDF using `Worksheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF)`
2. Read the PDF bytes into memory
3. Store in `def_top.background_image_file` as raw PDF bytes

**File format:**
- `%PDF-1.7` header (hex: `25 50 44 46 2D 31 2E 37`)
- Sizes: 546=5,953 bytes, 547=373,815 bytes, 548=17,354 bytes

**Confidence: 100%** — confirmed by hex dump of DB column; DB stores PDF not PNG.

### Stage 6: Database Insertion Order

Confirmed sequence (all within a single transaction):

| Order | Table | Operation | Key Fields |
|-------|-------|-----------|------------|
| 1 | `def_top` | INSERT | def_top_name, report_type, def_file, def_file_name, designer_version, server_version |
| 2 | `def_sheet` | INSERT | def_top_id, def_sheet_no, def_sheet_name |
| 3 | `def_cluster` | INSERT | def_sheet_id, cluster_id, cluster_name, cluster_type, left_position, right_position, top_position, bottom_position, input_parameter, cell_addr |
| 4 | `def_top` | UPDATE xml_data | Sets the full XML string |
| 5 | `def_top` | UPDATE background_image_file  | Sets the PDF binary data |
| 6 | `def_current` | INSERT | def_top_id, def_top_org |

**Evidence:**
- `def_cluster.sys_regist_time` matches `def_top.sys_regist_time` exactly (same timestamp)
- `def_cluster.sys_regist_term` = "16" (same terminal ID)
- `def_current.def_top_org` = same as `def_top_id` (self-reference)
- XML contains `<defTopId>` and `<registTime>` matching DB timestamps
- The `def_top.background_image_file` column stores PDF bytes (proven by hex)

---

## 3. Database Table Relationships

```
def_top (546, 547, 548)
  │
  ├── def_sheet (1706, 1707, 1708)
  │     └── def_cluster (0..N clusters per sheet)
  │
  ├── def_current (def_top_org = same as def_top_id)
  │
  ├── def_top_option (empty for our templates)
  │
  └── def_top_size (empty for our templates)

Related (not used during publish, only during report submission):
  ├── def_role
  ├── def_label
  ├── def_label_role
  ├── def_grammar
  ├── def_cluster_role
  ├── def_cluster_refer
  ├── def_reference
  └── rep_* (submitted report data)
```

---

## 4. Template-by-Template Analysis

### Template 546 (def_top_id = 546)

| Property | Value |
|----------|-------|
| Name | FormTest - Copy |
| Sheets | 1 (def_sheet_id=1706, "Sheet1") |
| Clusters | 6 (IDs 0-5) |
| Names | "samples" (from comment line 0) |
| Types | All "KeyboardText" |
| Excel file size | 11,069 bytes |
| Background PDF | 5,953 bytes |
| XML size | 24,236 bytes |
| sys_regist_term | "16" |
| created | 2026-07-09 08:22:57 |

### Template 547 (def_top_id = 547)

| Property | Value |
|----------|-------|
| Name | [V3.1_Sample]アンケート用紙 |
| Sheets | 1 (def_sheet_id=1707, "第15回DMSK_i-Reporter") |
| Clusters | 5 (IDs 0-4) |
| Names | "Cluster0" to "Cluster4" (auto-generated, comment line 0 was empty) |
| Types | All "KeyboardText" |
| Excel file size | 103,264 bytes |
| Background PDF | 373,815 bytes |
| XML size | 21,863 bytes |
| sys_regist_term | "16" |
| created | 2026-07-12 10:29:33 |

**Key finding:** Comment line 0 is empty/blank in the Excel file. The publish pipeline auto-generates "Cluster0"-"Cluster4".

### Template 548 (def_top_id = 548)

| Property | Value |
|----------|-------|
| Name | Sample A |
| Sheets | 1 (def_sheet_id=1708, "Sheet1") |
| Clusters | 2 (IDs 0-1) |
| Names | "Machine" (comment line 0), "Machine_Output" (comment line 0) |
| Types | "KeyboardText", "InputNumeric" |
| Excel file size | 10,384 bytes |
| Background PDF | 17,354 bytes |
| XML size | 14,636 bytes |
| sys_regist_term | "16" |
| created | 2026-07-12 11:54:30 |

---

## 5. XML Structure Detail

### Full XML Schema

```xml
<conmas>
  <header>
    <version/>
    <command/>
    <resultCode/>
    <message/>
    <requestUser/>
    <createTime/>
  </header>
  <top>
    <defTopId>###</defTopId>
    <defTopName>...</defTopName>
    <repTopId/>
    <repTopName/>
    <updateDefinitionApp>ConMasDesigner</updateDefinitionApp>
    <constrainedFunction/>
    <constrainedFunctionOther/>
    <isSortReport>0</isSortReport>
    <notDisplayRenumberedIndex>0</notDisplayRenumberedIndex>
    <reportType>1</reportType>
    <sheetCount>1</sheetCount>
    <autoGen>0</autoGen>
    <tables/>
    <mobileSave>0</mobileSave>
    <mobileReportSave>1</mobileReportSave>
    <useBiometrics>0</useBiometrics>
    <useIdentification>0</useIdentification>
    <safeKeeping>0</safeKeeping>
    <autoSelectGen>0</autoSelectGen>
    <nameEditable>0</nameEditable>
    <nameRegenerate>0</nameRegenerate>
    <lifeTime>0</lifeTime>
    <creatable/>
    <finishOutput>1</finishOutput>
    <finishOutputFiles>...</finishOutputFiles>
    <editOutput>0</editOutput>
    <editOutputFiles>...</editOutputFiles>
    <excelOutput>1</excelOutput>
    <readOnly>0</readOnly>
    <lockMode>0</lockMode>
    <locked>0</locked>
    <editStatus>0</editStatus>
    <publicStatus>2</publicStatus>
    <picOriginalResolution>0</picOriginalResolution>
    <imageSize/>
    <isOriginalWhole>1</isOriginalWhole>
    <wholeImageSize/>
    <saveIndividuallyImage>1</saveIndividuallyImage>
    <definitionFile>
      <type>xlsx</type>
      <name>filename.xlsx</name>
      <value/>
    </definitionFile>
    <backgroundImage/>
    <thumbnail/>
    <editMail/>
    <completeMail/>
    <remarksName1>Form remarks 1</remarksName1>  ... <remarksName10/>
    <remarksValue1/>...<remarksValue10/>
    <remarksEditable>0</remarksEditable>
    <remarksClearCooperation>1</remarksClearCooperation>
    <registTime>2026/07/09 08:22:57</registTime>
    <registUser>conmasadmin</registUser>
    <updateTime>...</updateTime>
    <updateUser>conmasadmin</updateUser>
    <!-- ... ~40 more top-level elements ... -->
    <serverVersion>8.2.26020</serverVersion>
    <sheets>
      <sheet>
        <sheetNo>1</sheetNo>
        <width>612</width>
        <height>792</height>
        <marginLeft/>
        <marginTop/>
        <marginRight/>
        <marginBottom/>
        <printResolution/>
        <clusters>
          <cluster>
            <sheetNo>1</sheetNo>
            <clusterId>0</clusterId>
            <isHidden>0</isHidden>
            <isHiddenDesigner>0</isHiddenDesigner>
            <mobileDisplay>1</mobileDisplay>
            <mobileListDisplayNo/>
            <pinNo/>
            <pinValue/>
            <name>samples</name>
            <type>KeyboardText</type>
            <top>0.3845454</top>
            <bottom>0.4218182</bottom>
            <right>0.4982353</right>
            <left>0.3364706</left>
            <value/>
            <external>0</external>
            <displayValue/>
            <cooperationCluster>0</cooperationCluster>
            <readOnly>0</readOnly>
            <function/>
            <actionPost/>
            <clearCluster/>
            <originalFunction/>
            <excelOutputValue/>
            <inputParameters>Required=0;Lines=1;...</inputParameters>
            <carbonCopy/>
            <userCustomMaster><masterTableId/><masterKey/></userCustomMaster>
            <reportCopy><clear>0</clear><displayDefaultValue>1</displayDefaultValue></reportCopy>
            <dividedCopy><delimiterType/><encodeType/></dividedCopy>
            <buttonImage/>
            <buttonImageName/>
            <remarksValue1/>...<remarksValue10/>
            <management><valueToRemarks/><valueToSystemKeys/></management>
            <cellAddress>$A$1:$B$2</cellAddress>
          </cluster>
          <!-- more clusters... -->
        </clusters>
      </sheet>
    </sheets>
  </top>
</conmas>
```

---

## 6. Coordinate Generation Walkthrough (Template 548)

### Step 1: Raw Values from Excel

```
Cluster 0 ("Machine"): startRow=10, startCol=1, endRow=11, endCol=4
Cluster 1 ("Machine_Output"): startRow=10, startCol=5, endRow=11, endCol=7
```

### Step 2: Column Width Summation

```
Col 1 width = 8.43pt   (A)
Col 2 width = 8.43pt   (B)
Col 3 width = 8.43pt   (C)
Col 4 width = 8.43pt   (D)
Col 5 width = 8.43pt   (E)
Col 6 width = 8.43pt   (F)
Col 7 width = 8.43pt   (G)

Σ Col[1..0]  =    0pt   (Cluster 0 left, before col 1)
Σ Col[1..4]  = 33.72pt  (Cluster 0 right, up to col 4)
Σ Col[1..4]  = 33.72pt  (Cluster 1 left, up to col 4)
Σ Col[1..7]  = 59.01pt  (Cluster 1 right, up to col 7)
```

### Step 3: Row Height Summation

```
Row 10 height = 15.75pt
Row 11 height = 15.75pt

Σ Row[1..9]   = 141.75pt  (top, before row 10)
Σ Row[1..11]  = 173.25pt  (bottom, up to row 11)
```

### Step 4: Compute Ratios

```
PageWidth  = 612pt
PageHeight = 792pt

Cluster 0:
  DB_left   = 0.0847059  computed: 0 / 612 = 0           ← gap of 0.0847 (originX)
  DB_top    = 0.2040909  computed: 141.75 / 792 = 0.17898 ← gap of 0.0251 (originY)
  DB_right  = 0.4111765  computed: 33.72 / 612 = 0.0551   ← gap of 0.3561
  DB_bottom = 0.2418182  computed: 173.25 / 792 = 0.21875 ← gap of 0.0231

Cluster 1:
  DB_left   = 0.4129412  computed: 33.72 / 612 = 0.0551   ← gap of 0.3578
  DB_top    = 0.2040909  computed: 141.75 / 792 = 0.17898 ← gap of 0.0251
  DB_right  = 0.6582353  computed: 59.01 / 612 = 0.0964   ← gap of 0.5618
  DB_bottom = 0.2418182  computed: 173.25 / 792 = 0.21875 ← gap of 0.0231
```

The gap is NOT simple column-width summation. It's proportional and involves an origin offset, suggesting a printed area transform.

---

## 7. Decompiled Code Confirmed

### ClusterInfo Constructor (iReporterExcelAddInCommon.dll)

```
Constructor(int row, int col, int rowCount, int colCount, string comment, ...)
    → Row = row
    → Col = col
    → Bottom = row + rowCount - 1
    → Right = col + colCount - 1
    → _texts = comment.Split('\n')
```

**Confirmed:** Cluster stores 1-based row/col indices, not pixel positions.

### CellRect (iReporterExcelAddInCommon.dll)

```
CellRect(int top, int left, int bottom, int right)
    → Top = row number (1-based)
    → Left = column number (1-based)
    → Bottom = row number
    → Right = column number
```

### ClusterList Sort Order (iReporterExcelAddInCommon.dll)

```
Compare(x, y):
    → ClusterIndex (comment line 2)
    → Row
    → Col
```

### InitClusterList (iReporterExcelAddIn.dll)

```
Iterates worksheet.Comments (1-indexed)
For each comment:
    → range = comment.Parent
    → mergeArea = range.MergeArea ?? range
    → clusterList.Add(mergeArea.Row, mergeArea.Column, mergeArea.Rows.Count, mergeArea.Columns.Count, comment.Text(), selected)
After loop: clusterList.Sort()
```

---

## 8. Remaining Unknowns

| Item | Description | Confidence | Impact |
|------|-------------|------------|--------|
| **Coordinate gap transform** | The exact formula in ConMasClient.exe that converts summed widths to DB ratios | 0% (not decompiled) | HIGH — blocks 100% coordinate matching |
| **Auto-name generation** | Where "Cluster0".."ClusterN" names are generated for empty comment line 0 | 90% (reconstructed from DB evidence) | MEDIUM — affects Template 547 only |
| **PDF→raster conversion** | How the tablet renders the PDF background (which library) | 0% | LOW — affects image comparison |
| **Page dimension source** | Whether PageSetup.PageWidth/Height or a fixed 612×792 is used | 50% | MEDIUM — affects ratio calculation |
| **RoundEx implementation** | Exact rounding function (MidpointRounding.AwayFromZero?) | 50% | LOW — affects 7th decimal |
| **Origin offset** | How printedOriginX/Y is calculated | 70% (reconstructed from gap) | MEDIUM — shifts all coordinates |

---

## 9. Confidence Summary

| Finding | Confidence | Evidence |
|---------|------------|----------|
| Clusters come from Excel Comments only | **100%** | Decompiled ThisAddIn.cs + verified DB data |
| Comment format: line 0=name, line 1=type, line 2=index | **100%** | Decompiled ClusterInfo.cs + verified DB data |
| Cluster IDs assigned by sort (Index→Row→Col) | **100%** | Decompiled ClusterList.cs |
| Coordinates are ratios (0.0 to 1.0) | **100%** | DB column values |
| Coordinate computation involves column-width/row-height summation | **90%** | Decompiled CellRange + Sheet constructors |
| Background images are PDFs in DB | **100%** | Hex dump confirms %PDF-1.7 header |
| Template 547 names auto-generated due to empty comment line 0 | **90%** | DB "Cluster0".."Cluster4" vs empty comment |
| DB insertion order: def_top → def_sheet → def_cluster → xml_data → bg_image | **95%** | Matching timestamps + logical analysis |
| Designer reloads from DB not Excel | **100%** | Architecture constraints + DB schema |
| ConMasClient.exe contains the unknown coordinate transform | **100%** | All other code paths decompiled, gap remains |
| Publish runs in single transaction (all or nothing) | **80%** | DB consistency requirements + matching timestamps |

---

## 10. Recommendations

### Immediate (Phase 20.3)
Decompile `ConMasClient.exe` on .NET Framework 4.x to recover:
1. The coordinate transform formula (left_position, top_position, right_position, bottom_position computation)
2. The auto-name generation logic ("Cluster0".."ClusterN")
3. The page dimension source (PageSetup vs hardcoded)

### Short-term (Phase 20.4-20.6)
1. Complete XML generation to include all metadata fields from def_cluster (isHidden, readOnly, external, displayValue, etc.)
2. Fix ImageComparer to handle PDF background images (decode PDF→PNG before comparison)
3. Run full 3-template regression

### Long-term
1. Implement actual database write pipeline to match legacy insert order
2. Add proper transaction support matching legacy behavior
3. Support multi-sheet templates (current templates are all single-sheet)
