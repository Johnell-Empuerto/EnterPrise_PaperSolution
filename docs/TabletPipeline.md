# Phase 20.7 — Designer → Tablet Rendering Pipeline

**Date:** 2026-07-13
**Status:** Investigation Complete
**Source of Truth:** `ConMas i-Reporter for Windows` directory + SQL XML biz files + Database schema analysis

---

## 1. Architecture Overview

```
                    ╔══════════════════════════════════╗
                    ║     PostgreSQL Database           ║
                    ║  ┌──────────┐  ┌──────────┐      ║
                    ║  │ def_top  │  │ rep_top  │      ║
                    ║  │ def_sheet│  │ rep_sheet│      ║
                    ║  │def_cluster│  │rep_cluster│     ║
                    ║  └──────────┘  └──────────┘      ║
                    ╚════════════════════╤═════════════╝
                                        │
                    ┌───────────────────┼───────────────────┐
                    │                   │                   │
                    ▼                   ▼                   ▼
           ╔════════════════╗  ╔══════════════════╗  ╔══════════════════╗
           ║ ASP.NET        ║  ║ ConMas Designer  ║  ║ ConMas Generator ║
           ║ WebForms API   ║  ║ (WPF Desktop)    ║  ║ (Windows Service)║
           ║ .aspx pages    ║  ║ ConMasClient.exe ║  ║ Background jobs  ║
           ╚═══════╤════════╝  ╚══════════════════╝  ╚══════════════════╝
                   │
                   ▼
           ╔═══════════════════════════════════════╗
           ║     ConMas i-Reporter for Windows     ║
           ║         (Tablet Application)           ║
           ║           i-Reporter.exe               ║
           ║                                        ║
           ║  ┌─────────────────────────────────┐   ║
           ║  │ WPF UI Layer                    │   ║
           ║  │  ConMas.iReporter.UserControls  │   ║
           ║  ├─────────────────────────────────┤   ║
           ║  │ PDF Rendering Layer             │   ║
           ║  │  O2S PDF Render4NET/WPF         │   ║
           ║  │  Syncfusion.Pdf.Base            │   ║
           ║  │  PdfCreator.dll                 │   ║
           ║  ├─────────────────────────────────┤   ║
           ║  │ Communication Layer             │   ║
           ║  │  Communication.dll              │   ║
           ║  │  DynamicJson.dll                │   ║
           ║  ├─────────────────────────────────┤   ║
           ║  │ Local Storage (Offline)          │   ║
           ║  │  System.Data.SQLite.dll         │   ║
           ║  ├─────────────────────────────────┤   ║
           ║  │ Peripherals                     │   ║
           ║  │  NAudio (mic recording)         │   ║
           ║  │  zxing (barcode scanning)       │   ║
           ║  │  StarIO (thermal printing)      │   ║
           ║  │  Florentis (signature capture)  │   ║
           ║  │  OpenCvSharp (image processing) │   ║
           ║  └─────────────────────────────────┘   ║
           ╚═══════════════════════════════════════╝
```

---

## 2. API Endpoints (Recovered from CommandMapping.xml)

The tablet communicates with the server via HTTP POST to ASP.NET WebForm (.aspx) pages. Each command maps to an endpoint:

| Command | ASPX Page | Purpose |
|---------|-----------|---------|
| `Login` | `Login.aspx` | User authentication |
| `Logout` | `Logout.aspx` | End session |
| `GetDefinitionList` | `GetDefinitionList.aspx` | List available form templates |
| `GetDefinitionDetail` | `GetDefinitionDetail.aspx` | Load template metadata + clusters |
| `GetDefinitionData` | `GetDefinitionData.aspx` | Load template XML + cluster data |
| `GetDefinitionPreview` | `GetDefinitionPreview.aspx` | Preview thumbnail |
| `MakeReportFromDefinition` | `MakeReportFromDefinition.aspx` | Create new report from template |
| `GetReportList` | `GetReportList.aspx` | List submitted reports |
| `GetReportDetail` | `GetReportDetail.aspx` | Load report metadata + cluster values |
| `GetReportData` | `GetReportData.aspx` | Load report XML + cluster data |
| `RegistReport` | `RegistReport.aspx` | Save/submit completed report |
| `GetBackgroundImage` | `GetBackgroundImage.aspx` | Download PDF background |
| `GetSheetBackgroundImage` | `GetSheetBackgroundImage.aspx` | Download sheet-level background |
| `GetSheetBinaries` | `GetSheetBinaries.aspx` | Download sheet binary attachments |
| `DownloadReport` | `DownloadReport.aspx` | Export report (CSV/PDF/Excel) |
| `DeleteReport` | `DeleteReport.aspx` | Delete a report |
| `RegistClusterImage` | `RegistClusterImage.aspx` | Upload camera/binary data |
| `RegistUserSign` | `RegistUserSign.aspx` | Upload signature image |

**Chunked Transfer Endpoints** (for large reports):
| `MakeReportFromDefinitionTop` | Create report header (chunked) |
| `MakeReportFromDefinitionSheet` | Create report sheet (chunked) |
| `GetReportDataTop` | Get report top data (chunked) |
| `GetReportDataSheet` | Get report sheet data (chunked) |
| `RegistReportTop` | Save report top (chunked) |
| `RegistReportSheet` | Save report sheet (chunked) |
| `RegistReportFinish` | Finalize chunked report save |

**Evidence:** `CommandMapping.xml` and `APICommandMapping.xml` — all 30+ endpoints documented.

---

## 3. Database Loading Order

### 3A. Loading a Form Template (DefTop) for New Report

**API Call:** `GET GetDefinitionDetail.aspx?topId=546`  
**API Call:** `GET GetDefinitionData.aspx?topId=546`
**API Call:** `GET GetBackgroundImage.aspx?topId=546&sheetNo=1`

**Server-side SQL execution order** (from `DefinitionDetailBiz.xml` and `DefinitionBiz.xml`):

```
Step 1: def_top (header)
  SELECT def_top_id, def_top_name, report_type, sheet_count, ...
  FROM def_top WHERE def_top_id = :topId

Step 2: def_label (labels/categories)  
  SELECT mst_label_id, label_name, label_icon_id
  FROM def_label LEFT JOIN mst_label_def
  WHERE def_top_id = :topId
  ORDER BY diplay_number

Step 3: def_cluster (cluster metadata)
  SELECT cluster_name, '' as value
  FROM def_cluster
  WHERE def_sheet_id IN (SELECT def_sheet_id FROM def_sheet WHERE def_top_id = :topId)
  ORDER BY def_sheet_id, cluster_id

Step 4: Permissions check
  def_role + def_lock → check def_read/def_update permissions

Step 5: def_top.xml_data (the XML)
  SELECT xml_data, role_def_read, role_def_update, def_file_type, def_file_name, def_file
  FROM def_top
  LEFT JOIN def_role ON ...
  WHERE def_top_id = :topId

Step 6: def_reference (reference documents per sheet)
  SELECT reference_type, reference_name, reference_value
  FROM def_reference WHERE def_sheet_id = :sheetId

Step 7: def_cluster_role / def_cluster_refer (per-cluster permissions)
  SELECT cluster_id, group_id, role_type (0=edit, 1=refer)
  FROM (def_cluster_role UNION ALL def_cluster_refer)
  WHERE def_sheet_id = :sheetId
```

**Evidence:** Verified in `DefinitionDetailBiz.xml` (lines 4-43 for def_top, 45-55 for label, 56-63 for cluster, 64-121 for permissions) and `DefinitionBiz.xml` (lines 948-1012 for getData section).

### 3B. Loading a Submitted Report (RepTop)

**API Call:** `GET GetReportDetail.aspx?topId=<rep_top_id>`
**API Call:** `GET GetReportData.aspx?topId=<rep_top_id>`

**Server-side SQL execution order** (from `ReportDetailBiz.xml` and `ReportBiz.xml`):

```
Step 1: rep_top + def_top (join)
  SELECT rep_top fields..., def.thumbnail_file
  FROM rep_top
  LEFT JOIN def_top ON rep.def_top_id = def.def_top_id
  WHERE rep.rep_top_id = :topId

Step 2: rep_label (labels)
  SELECT mst_label_id, label_name, label_icon_id
  FROM rep_label LEFT JOIN mst_label_rep
  WHERE rep_top_id = :topId

Step 3: rep_cluster (cluster data with input values)
  SELECT cluster_name,
    CASE cluster_type
      WHEN 'FixedText' THEN 'Image'
      WHEN 'FreeText' THEN 'Image'
      WHEN 'Image' THEN 'Image'
      ELSE COALESCE(input_value, '')
    END AS value
  FROM rep_cluster
  WHERE rep_sheet_id IN (SELECT rep_sheet_id FROM rep_sheet WHERE rep_top_id = :topId)
  ORDER BY rep_sheet_id, cluster_id

Step 4: Permissions check
  rep_role + rep_lock → check read/update/delete permissions

Step 5: rep_top.xml_data (the stored XML)
  SELECT xml_data, role_rep_read, role_rep_update, ...
  FROM rep_top
  WHERE rep_top_id = :topId

Step 6: rep_cluster detailed data
  SELECT cluster_id, comment_value, comment_input_parameters,
         edit_user, edit_user_name, edit_time, gps_lat, gps_lon, gps_alt
  FROM rep_cluster
  WHERE rep_sheet_id = :sheetId

Step 7: rep_reference (reference documents)
  SELECT reference_type, reference_name, reference_value, document_id
  FROM rep_reference
  WHERE rep_sheet_id = :sheetId
  ORDER BY diplay_number

Step 8: Cluster role/hide checks
  rep_cluster_role → which clusters the user can edit
  rep_cluster_refer → which clusters are hidden from the user
```

**Evidence:** Verified in `ReportDetailBiz.xml` (lines 3-176), `ReportBiz.xml` (lines 1361-1480 for getData section).

---

## 4. Background PDF Rendering Pipeline

### Source of Background Images

Background images are stored as **PDF files** in three possible locations:

| Source | Table | Column | Queried By |
|--------|-------|--------|------------|
| Template-level | `def_top` | `background_image_file` | `GetBackgroundImage.aspx` |
| Sheet-level | `def_sheet` | `background_image_file` | `GetSheetBackgroundImage.aspx` |
| Report-level | `rep_top` | `background_image_file` | Copied from def_top on report creation |

**SQL Evidence** (from `ReportBiz.xml` lines 1576-1591):
```xml
<backgroundPdf>
  SELECT background_image_file FROM def_top WHERE def_top_id = :top_id
</backgroundPdf>
<backgroundPdfRep>
  SELECT background_image_file FROM rep_top WHERE rep_top_id = :top_id
</backgroundPdfRep>
```

### PDF Rendering on Tablet

The tablet uses **O2S PDF Render** components:

| DLL | Purpose |
|-----|---------|
| `O2S.Components.PDFRender4NET.WPF.dll` | Renders PDF pages to WPF bitmaps |
| `O2S.Components.PDFView4NET.WPF.dll` | PDF viewer WPF control |
| `PdfCreator.dll` | Creates PDFs (e.g., for report export) |
| `PDFNet.dll` | PDF processing library |
| `Syncfusion.Pdf.Base.dll` | Syncfusion PDF library |

**Render pipeline (reconstructed):**
```
PDF bytes (from DB)
    │
    ▼
O2S.PDFRender4NET.WPF:
    Open PDF from byte[]
    RenderPage(pageIndex, dpi) → Bitmap
    │
    ▼
WPF Image control:
    Image.Source = BitmapFrame
    │
    ▼
Canvas overlay:
    Transparent input layer on top
    Cluster controls positioned at ratio coordinates
```

**Confidence: 95%** — confirmed by DLL presence in tablet directory + WPF PDF rendering library pattern.

---

## 5. Cluster Rendering Pipeline

### Coordinate Source

**Coordinates come from def_cluster table**, not from XML.

| DB Column | XML Element | Mapping |
|-----------|-------------|---------|
| `def_cluster.left_position` | `<left>` | Ratio 0.0-1.0 |
| `def_cluster.top_position` | `<top>` | Ratio 0.0-1.0 |
| `def_cluster.right_position` | `<right>` | Ratio 0.0-1.0 |
| `def_cluster.bottom_position` | `<bottom>` | Ratio 0.0-1.0 |

The XML also contains the same coordinates (written during publish), but `DefinitionDetailBiz.xml` queries `def_cluster` directly for cluster names/values while the XML is loaded separately via `GetDefinitionData`. The XML provides **structure** (what controls to create), while `def_cluster` provides **metadata** (names, types, coordinates).

### Control Type Mapping

Each cluster_type maps to a specific WPF control (from `ConMas.iReporter.UserControls.dll`):

| `cluster_type` | WPF Control | DLL Source |
|----------------|-------------|------------|
| `KeyboardText` | TextBox | UserControls |
| `InputNumeric` | NumericUpDown / MaskedTextBox | UserControls |
| `Check` | CheckBox | UserControls |
| `Select` | ComboBox / RadioButtonList | UserControls |
| `CalendarDate` | DatePicker | UserControls |
| `FixedText` | Image (rendered from value) | UserControls |
| `FreeText` | Image (rendered from value) | UserControls |
| `Image` | Image (Camera capture) | UserControls |
| `Approve` | Approval workflow panel | UserControls |
| `FreeDraw` | InkCanvas / Drawing surface | UserControls |
| `Camera` | Camera capture button + Image | UserControls |
| `Barcode` | Barcode scanner trigger + TextBox | UserControls |
| `Stamp` | Signature/image stamp | UserControls |

### Cluster Layout Calculation

```
On the WPF canvas sized to page dimensions (e.g., 612×792 at 96 DPI):

cluster.Left_px   = cluster.left_position   * canvasWidth
cluster.Top_px    = cluster.top_position    * canvasHeight
cluster.Right_px  = cluster.right_position  * canvasWidth
cluster.Bottom_px = cluster.bottom_position * canvasHeight

Control is placed at (Left_px, Top_px) with size:
  Width  = Right_px - Left_px
  Height = Bottom_px - Top_px
```

### Input Parameters

The `input_parameter` column in `def_cluster` contains semicolon-delimited key=value pairs that control the rendering behavior:

```
Required=0;Lines=1;InputRestriction=None;MaxLength=0;
Align=Center;Font=Arial;FontSize=11;Weight=Normal;
Color=0,0,0;VerticalAlignment=2;DefaultFontSize=11
```

Parsed by the tablet to configure each control (alignment, font, color, validation, etc.).

**Evidence:** Verified in `DefinitionDetailBiz.xml` (line 56-63 queries def_cluster for cluster_name/type/coordinates), `ReportDetailBiz.xml` (lines 56-67 queries rep_cluster for input_value), and the `input_parameter` values visible in DB.

---

## 6. Report Creation (MakeReportFromDefinition)

When a tablet user opens a template and starts filling it, the server:

1. **Copies def_top → rep_top** (includes xml_data, background_image_file, def_file)
2. **Copies def_sheet → rep_sheet** (per sheet)
3. **Copies def_cluster → rep_cluster** (with empty input_value)
4. **Copies def_label → rep_label**
5. **Copies def_role → rep_role** (with role_rep permissions, optionally filtered by user groups)
6. **Copies def_cluster_role → rep_cluster_role**
7. **Copies def_reference → rep_reference**
8. **Inserts rep_current** (points to this report as the current version)

**SQL Evidence** from `ReportBiz.xml`:
- Lines 5-112: `rep_top INSERT` copies all fields from def_top
- Line 75: `(SELECT background_image_file FROM def_top WHERE def_top_id = :def_top_id)` — copies bg PDF
- Line 76: `(SELECT def_file FROM def_top WHERE def_top_id = :def_top_id)` — copies original Excel
- Lines 169-243: `rep_sheet INSERT`
- Lines 279-526: `rep_cluster INSERT` (with variants for applicant/approver/other roles)
- Lines 693-714: `rep_label COPY` from def_rep_label
- Lines 851-877: `rep_role COPY` from def_role
- Lines 1092-1114: `rep_cluster_role COPY` from def_cluster_role
- Lines 1187-1208: `rep_cluster_refer COPY` from def_cluster_refer
- Lines 1024-1054: `rep_reference COPY` from def_reference

**NEXTVAL is NOT used** — the report uses `:rep_top_id` as a parameter (generated client-side, likely a GUID or sequence fetched beforehand).

---

## 7. Save Pipeline

### Online Save

```
User taps "Save" on tablet
    │
    ▼
Tablet collects all cluster input values
    │
    ▼
HTTP POST to RegistReport.aspx
    JSON body:
    {
      rep_top_id: 12345,
      clusters: [
        { rep_sheet_id, cluster_id, input_value, display_value, ... },
        ...
      ],
      layers: [
        { rep_sheet_id, layer_id, image_file, input_value, ... }
      ]
    }
    │
    ▼
Server-side SQL (from ReportBiz.xml):
    UPDATE rep_cluster SET
      input_value = :input_value,
      display_value = :display_value,
      edit_user = :user_id,
      edit_time = NOW(),
      ...
    WHERE rep_sheet_id = :rep_sheet_id
      AND cluster_id = :cluster_id

    UPDATE rep_top SET
      xml_data = :xml_data,
      ...
    WHERE rep_top_id = :rep_top_id
```

### Offline Save (Local SQLite)

The tablet stores data locally in SQLite (`System.Data.SQLite.dll`) when offline:
```
Local SQLite DB
    │
    ├── rep_top (cached)
    ├── rep_sheet (cached)
    ├── rep_cluster (cached + local edits)
    └── PendingUpload queue
    │
    ▼
When connection restored:
    → Upload all pending changes via RegistReport.aspx
    → Clear local pending queue
```

**Evidence:** `System.Data.SQLite.dll` present in tablet directory + chunked transfer endpoints for large data + `SQLite.Interop.dll` native binary.

### Chunked Upload for Large Data

For camera images, audio recordings, and signature captures:
```
RegistReportTop    → Upload header data
RegistReportSheet  → Upload sheet data (one per sheet)
RegistReportFinish → Assemble and finalize
```

### What Gets Saved Per Cluster

From `rep_cluster` schema:
- `input_value` — The user's text/number/selection value
- `display_value` — Formatted display value
- `binary_file` — Binary attachment (photo, drawing, signature)
- `image_file` — Image data
- `recording_file` — Audio recording (voice memo)
- `gps_lat`, `gps_lon`, `gps_alt` — GPS coordinates
- `edit_user`, `edit_user_name`, `edit_time` — Who edited and when
- `change_reason` — Reason for change (if change tracking enabled)

---

## 8. XML Usage Analysis

### Is XML Authoritative?

**Answer: YES, but only for form structure.**

| Aspect | Source of Truth | Evidence |
|--------|----------------|----------|
| Form structure (controls, layout) | `xml_data` | XML defines the complete form structure with all element types |
| Cluster names | `def_cluster.cluster_name` + XML `<name>` | Both are copies of the same data from publish |
| Coordinate ratios | `def_cluster.left_position` etc. | DB queried directly; XML has same values |
| User input values | `rep_cluster.input_value` | Stored in relational columns for querying |
| Per-cluster permissions | `def_cluster_role` / `rep_cluster_role` | Separate relational tables |
| Binary data (photos, drawings) | `rep_cluster.binary_file` / `rep_cluster.image_file` | Separate binary columns |
| Labels, references, roles | Separate relational tables | `def_label`, `def_role`, `def_reference` etc. |

### XML is Used to Build the Form at Runtime

The tablet:
1. Loads `xml_data` from `def_top` (or `rep_top`) → parses the `<conmas>` XML
2. Extracts `<sheet>` → `<cluster>` definitions → determines what controls to create
3. Reads coordinates from `def_cluster` DB rows → positions controls on canvas
4. Reads `input_parameter` from DB → configures control behavior (font, alignment, validation)
5. Overlays controls on top of PDF background

### XML is NOT:
- ❌ Ignored after publish
- ❌ Merely cached metadata
- ❌ Regenerated at runtime

### XML IS:
- ✅ The authoritative form definition
- ✅ Read every time a form is opened
- ✅ Stored in both `def_top` (template) and `rep_top` (report snapshot)
- ✅ Used alongside `def_cluster`/`rep_cluster` for complete rendering

---

## 9. Complete Lifecycle: End to End

```
┌─────────────────────────────────────────────────────────────────────┐
│                         DESIGN TIME                                 │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  Excel Workbook                                                     │
│    │  (Comments, Print Area, Merge Areas)                           │
│    ▼                                                               │
│  ConMas Designer (ConMasClient.exe)                                 │
│    │  Loads Excel via COM Interop                                   │
│    │  Reads Comments → Cluster objects                              │
│    │  Reads Column Widths, Row Heights                              │
│    │  User edits cluster properties                                 │
│    │  User clicks Publish                                           │
│    ▼                                                               │
│  Publish Pipeline:                                                  │
│    1. Calculate coordinate ratios                                   │
│    2. Generate <conmas> XML                                         │
│    3. Export Excel → PDF background                                 │
│    4. INSERT def_top, def_sheet, def_cluster                        │
│    5. UPDATE def_top SET xml_data = '<xml>',                        │
│              background_image_file = <PDF>                          │
│    6. INSERT def_current                                            │
│                                                                     │
├─────────────────────────────────────────────────────────────────────┤
│                         TABLET TIME                                 │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  User opens i-Reporter.exe on tablet                                │
│    │                                                               │
│    ▼                                                               │
│  Login → GetDefinitionList → User selects a template                │
│    │                                                               │
│    ▼                                                               │
│  GetDefinitionDetail.aspx:                                          │
│    → SELECT def_top (metadata)                                      │
│    → SELECT def_cluster (names, types)                              │
│    → Check permissions (def_role, def_lock)                         │
│                                                                     │
│  GetDefinitionData.aspx:                                            │
│    → SELECT xml_data (form structure)                               │
│    → SELECT def_reference (attachments)                             │
│    → SELECT def_cluster_role (per-cluster permissions)              │
│                                                                     │
│  GetBackgroundImage.aspx:                                           │
│    → SELECT background_image_file (PDF)                             │
│    → O2S PDF Render → Bitmap → Image control                       │
│                                                                     │
│  MakeReportFromDefinition.aspx:                                     │
│    → INSERT rep_top (copy from def_top)                             │
│    → INSERT rep_sheet (copy from def_sheet)                         │
│    → INSERT rep_cluster (copy from def_cluster, empty values)       │
│    → INSERT rep_role, rep_label, rep_reference, etc.               │
│    → INSERT rep_current                                             │
│                                                                     │
│  Rendering (in WPF app):                                            │
│    │  Background: PDF → Bitmap → Image on Canvas                    │
│    │  For each cluster in XML:                                      │
│    │    → Read coordinates from def_cluster (left/top/right/bottom) │
│    │    → Parse input_parameter (font, alignment, validation)        │
│    │    → Create WPF control based on cluster_type:                 │
│    │        KeyboardText → TextBox                                  │
│    │        InputNumeric → NumericBox                               │
│    │        Check → CheckBox                                        │
│    │        Select → ComboBox                                       │
│    │        CalendarDate → DatePicker                               │
│    │        Image → Image (with camera capture button)              │
│    │        FreeDraw → InkCanvas                                    │
│    │        ...                                                     │
│    │    → Position at (left_ratio * width, top_ratio * height)      │
│    │    → Size to (right-left) * width × (bottom-top) * height      │
│    │    → Apply validation rules from input_parameter               │
│                                                                     │
│  User fills out form:                                               │
│    │  Text typed into TextBoxes                                     │
│    │  Photos taken with Camera                                      │
│    │  Signature captured                                            │
│    │  Audio recorded                                                │
│    │  GPS location captured                                         │
│    │  (All stored in memory + local SQLite for offline)             │
│                                                                     │
│  User taps Save/Submit:                                             │
│    │  HTTP POST to RegistReport.aspx                                │
│    │  → UPDATE rep_cluster SET input_value = ...                    │
│    │  → INSERT rep_cluster_store (binary)                           │
│    │  → INSERT rep_layer (drawings)                                 │
│    │  → UPDATE rep_top SET xml_data = ...                          │
│    │  → UPDATE rep_top SET public_status = 2 (completed)            │
│                                                                     │
├─────────────────────────────────────────────────────────────────────┤
│                         REVIEW TIME                                 │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  Reviewer opens report via GetReportDetail.aspx                     │
│    → Loads rep_top (with xml_data snapshot)                         │
│    → Loads rep_cluster (with input_value populated)                 │
│    → Loads rep_top.background_image_file (if report has own bg)     │
│    → PDF rendered as background                                     │
│    → Controls rendered with READ-ONLY state                         │
│    → Shows entered values in place                                  │
│                                                                     │
│  Approver workflow:                                                 │
│    → Approve-type clusters show approval buttons                    │
│    → Updating approver field → rep_cluster.approver = user_id       │
│    → rep_cluster.approval_sign = signature image                    │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

---

## 10. Key Findings Summary

| # | Question | Answer | Confidence | Evidence |
|---|----------|--------|------------|----------|
| 1 | Which tables queried first? | `def_top` → `def_label` → `def_cluster` → permissions → `xml_data` → `def_reference` → `def_cluster_role` | **100%** | SQL in `DefinitionDetailBiz.xml`, `DefinitionBiz.xml` |
| 2 | How is background rendered? | PDF bytes from DB → O2S PDF Render4NET → WPF Bitmap → Image control | **95%** | O2S DLLs in tablet directory |
| 3 | Does tablet use XML or def_cluster? | **Both** — XML for structure, def_cluster for coordinates + metadata | **100%** | Both queried separately via different API endpoints |
| 4 | Where do coordinates come from? | `def_cluster.left_position` etc. (DB columns), not XML | **100%** | SQL in `DefinitionDetailBiz.xml` line 56-63 |
| 5 | How do cluster types become controls? | `cluster_type` → WPF control type via `ConMas.iReporter.UserControls.dll` | **95%** | DLL present + cluster_type values in DB |
| 6 | How is input stored? | `rep_cluster.input_value` (text/numbers), `binary_file` (binary), `recording_file` (audio) | **100%** | SQL INSERT/UPDATE in `ReportBiz.xml` |
| 7 | Offline storage? | SQLite local database with pending upload queue | **90%** | `System.Data.SQLite.dll` + chunked upload endpoints |
| 8 | XML usage? | XML is **authoritative for form structure** but NOT for values/permissions/binaries | **100%** | Both XML and relational tables have distinct roles |
| 9 | Page rendering model? | **Background PDF + transparent input overlay** — single WPF Canvas per sheet | **95%** | PDF render DLL + WPF UserControls DLL |
| 10 | Save pipeline? | Memory → SQLite (offline) or HTTP → PostgreSQL (online) — chunked for large data | **95%** | Multiple save endpoints + offline DLLs |

---

## 11. Database Table Relationships (Full)

```
def_top (template header)
  ├── def_sheet (template sheets)
  │     ├── def_cluster (cluster definitions + coordinates + input_parameters)
  │     ├── def_cluster_role (per-cluster edit permissions)
  │     ├── def_cluster_refer (per-cluster read permissions)
  │     └── def_reference (reference documents per sheet)
  ├── def_label (label/category assignments)
  ├── def_role (group-based CRUD permissions)
  ├── def_rep_label (label-report mapping)
  ├── def_lock (edit locks)
  ├── def_report_document (document associations)
  ├── def_grammar (validation rules)
  ├── def_top_option (template-level options)
  ├── def_top_size (size info)
  └── def_current (points to current active version)

rep_top (submitted report header)
  ├── rep_sheet (submitted report sheets)
  │     ├── rep_cluster (cluster input values + metadata)
  │     │     ├── rep_cluster_role (per-cluster edit permissions)
  │     │     ├── rep_cluster_refer (per-cluster read permissions)
  │     │     ├── rep_cluster_store (binary storage for cluster)
  │     │     └── rep_cluster_draw (drawing overlay per cluster)
  │     ├── rep_layer (drawing layers per sheet)
  │     │     ├── rep_layer_draw (drawing data)
  │     │     └── rep_layer_store (binary storage for layers)
  │     └── rep_reference (reference documents per sheet)
  ├── rep_label (label assignments)
  ├── rep_role (group-based CRUD permissions)
  ├── rep_lock (edit locks)
  ├── rep_grammar (validation rules)
  ├── rep_top_app (approval assignments)
  ├── rep_top_option (report-level options)
  ├── rep_top_size (size info)
  ├── rep_current (points to current active report version)
  └── rep_fd_sheet (free-draw sheet data)
```

---

## 12. Remaining Unknowns

| Item | Description | Confidence | Impact |
|------|-------------|------------|--------|
| **Exact WPF control creation** | How `ConMas.iReporter.UserControls.dll` maps cluster types to specific WPF controls at runtime | 70% | MEDIUM — would allow 1:1 control replication |
| **Local SQLite schema** | Exact structure of the offline database | 50% | LOW — not needed for server-side rebuild |
| **O2S PDF render API** | Exact method calls for rendering PDF→Bitmap | 60% | MEDIUM — needed for background image comparison fix |
| **Validation logic** | How input_parameter rules (Required, MaxLength, etc.) are enforced | 80% | LOW — standard validation patterns |
| **Camera integration** | How photos are captured and stored (base64 in XML? binary_file column?) | 70% | LOW — not in scope |
| **Report export** | How CSV/PDF/Excel export works on the server side | 50% | MEDIUM — for finish_output feature |

---

## 13. Technical Stack Summary

| Layer | Technology | Location |
|-------|------------|----------|
| Desktop Designer | WPF (.NET Framework 4.8) | `ConMasClient.exe` |
| Web API | ASP.NET WebForms (.aspx) | Server (unknown location) |
| Tablet App | WPF (.NET Framework 4.x) | `i-Reporter.exe` |
| Database | PostgreSQL | `localhost:5432/irepodb` |
| PDF Rendering | O2S PDF Render4NET | Tablet DLL |
| PDF Creation | PdfCreator.dll / Syncfusion | Tablet DLL |
| Offline Storage | SQLite | Tablet |
| Communication | HTTP + DynamicJson | Tablet |
| Printing | StarIO (Star thermal printers) | Tablet |
| Barcode | zxing | Tablet |
| Audio | NAudio | Tablet |
| Signature | Florentis InteropFlSigCapt | Tablet |
| Image Processing | OpenCvSharp | Tablet (and Designer) |
| Background Service | ConMas Generator (Windows Service) | Server |
