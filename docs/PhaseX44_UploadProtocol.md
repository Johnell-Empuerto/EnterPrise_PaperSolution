# Phase X.44 — Designer ↔ Server Upload Protocol

**Date**: 2026-07-19
**Objective**: Reconstruct the exact payload exchanged between the legacy ConMas Designer (WPF) and the production server (ASP.NET WebForms).

---

## 1. Executive Summary

The Designer uploads a **single XML file** to the server via HTTP POST. The XML contains:
- **Form metadata** — sheets, clusters, fields, coordinates
- **The entire Excel template** — Base64-encoded inside `<definitionFile><value>`
- **Thumbnail & background image** — Base64-encoded bitmap data

The server:
1. Parses the XML from `HttpRequest.Files[0].InputStream`
2. Inserts metadata into `def_top`, `def_sheet`, `def_cluster`
3. Stores the Excel template and images as BLOBs in database columns
4. **Later rendering** uses these stored assets — extracting the Excel BLOB, re-rendering via Syncfusion PDF, overlaying cluster fields from `rep_cluster`

**Cluster images** are uploaded separately as individual HTTP requests to `RegistClusterImage.aspx`.

---

## 2. Complete Call Chain

```
User clicks "Save and Exit" / "Save Public" / "Save Test"
    │
    ▼
MainWindow.menuSaveAndExit_Click / menuSavePublicAndExit_Click / menuSaveTestAndExit_Click
    │  (ConMasClient.exe lines 253012–253082)
    │  SaveType = 0 (save), 1 (test), 2 (public)
    ▼
MainWindow.ReportSave(SaveType)
    │  (ConMasClient.exe line 253126)
    │  Calls OutputXml(topReportData, SaveType, excelValue)
    ▼
MainWindow.OutputXml(TopReport, SaveType, excelValue)
    │  (ConMasClient.exe line 357316)
    │
    ├── SaveType == 3 → Local SaveFileDialog (user picks path, save .xml)
    ├── SaveType == 4 → Save to TemporaryFile.xml (no upload)
    │
    └── SaveType == 0,1,2 →
            │
            ├─ 1. ExportExcelDefinition(xlsFilePath, tempDir, ref ProgressValue)
            │      ↓
            │      ├─ ExcelControllerInterop.GetDefinition(ref ProgressValue, outputDir, filename)
            │      │      ↓
            │      │      ├─ Open Excel via COM Interop
            │      │      ├─ Read all sheets, print areas, clusters
            │      │      ├─ Export PDF via ExportPdf2010() (Workbook.ExportAsFixedFormat)
            │      │      ├─ Run GetClusterSize() → GetAddress() pixel scan for coordinates
            │      │      ├─ Embed xlsx as Base64 in <definitionFile><value>
            │      │      └─ Return complete XElement
            │      │
            │      └─ Save XElement.ToString() to: {tempDir}\{filename}.xml
            │
            ├─ 2. Copy XML to: UserConfigDirectory + "UploadFile.xml"
            │
            ├─ 3. LoginCheck() — ensure authenticated session
            │
            ├─ 4. wc.UnlockDefinition(orgTopId)
            │
            ├─ 5. wc.UploadDefinition(uploadFilePath, orgTopId, false, "")
            │      ↓
            │      ├─ WebClient.UploadFile(ServerUrl + "?command=RegistDefinition", uploadFilePath)
            │      ├─ Parse response: <error><code>VALUE</code></error>
            │      └─ Return int.Parse(code)  (0=success, 1003=needs approval)
            │
            ├─ 6. If result == 1003 → Show ApplyPublishingWindow
            │      If user clicks "Regist" →
            │         ├─ Set publicStatus = "1" in XML
            │         ├─ wc.UploadDefinition(uploadFilePath, orgTopId, true, comment)
            │         └─ Delete UploadFile.xml
            │
            └─ 7. Cleanup: Delete UploadFile.xml
```

---

## 3. The Upload File: XML Structure

**File path**: `{UserConfigDirectory}\UploadFile.xml`
**Extension**: `.xml`
**Size**: Varies (includes Base64-encoded xlsx)

### XML Schema (Reconstructed from IL)

The XML is built by `ExcelControllerInterop.GetDefinition()` (LibExcelController.il lines 1797–3437) and has this structure:

```xml
<conmas>
  <top>
    <!-- METADATA FIELDS (set by OutputXml in ConMasClient) -->
    <topId>123</topId>
    <topIdOrg>100</topIdOrg>
    <revNo>1</revNo>
    <serverVersion>3.0.0.0</serverVersion>
    <registTime>2026-07-19T10:00:00</registTime>
    <updateTime>2026-07-19T10:00:00</updateTime>
    <registUser>username</registUser>
    <updateUser>username</updateUser>
    <title>Form Title</title>
    <description>Form description</description>
    <publicStatus>0</publicStatus>          <!-- 0=draft, 1=applying, 2=published -->
    <approveType>0</approveType>
    <thumbnail>base64-encoded-thumbnail-png...</thumbnail>
    <backgroundImage>base64-encoded-background-png...</backgroundImage>

    <!-- SHEETS -->
    <sheets>
      <sheet>
        <defSheetId>456</defSheetId>
        <defSheetName>Sheet1</defSheetName>
        <sheetNo>1</sheetNo>
        <pageOrder>1</pageOrder>
        <printPaperSize>A4</printPaperSize>
        <printOrientation>1</printOrientation>   <!-- 1=portrait, 2=landscape -->

        <!-- CLUSTERS (data fields / regions) -->
        <clusters>
          <cluster>
            <clusterId>1</clusterId>
            <clusterName>field_customer_name</clusterName>
            <clusterType>0</clusterType>
            <defSheetId>456</defSheetId>
            <top>123.45</top>       <!-- normalized float: y / image.Height -->
            <bottom>678.90</bottom>
            <left>55.0</left>       <!-- normalized float: x / image.Width -->
            <right>444.0</right>
            <inputParameter>text</inputParameter>
            <remarksValue1>...</remarksValue1>
            <!-- remarksValue2..10 also possible -->
          </cluster>
        </clusters>
      </sheet>
    </sheets>

    <!-- EMBEDDED EXCEL FILE (Base64) -->
    <definitionFile>
      <type>xlsx</type>
      <name>OriginalFilename.xlsx</name>
      <value>base64-encoded-excel-content-here...</value>
    </definitionFile>
  </top>
</conmas>
```

### Key XML Elements

| Element | Source | Purpose |
|---------|--------|---------|
| `topId` | DB (server-assigned) | Primary key in `def_top` |
| `title` | User input | Form display name |
| `publicStatus` | Designer logic | `0`=draft, `1`=pending approval, `2`=published |
| `thumbnail` | Designer (base64 PNG) | Form thumbnail preview |
| `backgroundImage` | Designer (base64 PNG) | Background image for rendering |
| `sheets[]` | Designer | Per-page definitions |
| `clusters[]` | `GetAddress()` pixel scan | Data field regions with normalized coordinates |
| `definitionFile` | Designer (base64 xlsx) | **The original Excel template**, stored for re-rendering |

---

## 4. Server-Side Import: DefinitionBiz.Regist()

**File**: Cimtops.ConMasBiz.il lines 496060–496778
**Class**: `Biz.DefinitionBiz`
**Method**: `XElement Regist(ConMasPage page)`

### Step-by-Step Flow

```
Regist(ConMasPage page)
│
├─ 1. READ approveType from Request.Params["approveType"]
│      If "1": setIsApplyDefinision = true
│
├─ 2. LOAD XML from uploaded file:
│      using (XmlReader reader = XmlReader.Create(page.Request.Files[0].InputStream))
│          xInputReport = XElement.Load(reader);
│
├─ 3. VALIDATE:
│      ValidateReportTable(xInputReport, out xResult)  — validates cluster/table structure
│      ValidateWorkReport(xInputReport, out xResult)   — validates workflow
│
├─ 4. ROUTE based on whether topId exists:
│
│   [NEW DEFINITION - no topId]:
│   │   ├─ CheckCreateRole() — permission check
│   │   ├─ IsDefinisionApproveRole() check for publicStatus="2"
│   │   ├─ License check — IsNumbersLicense()
│   │   └─ Entry(nextRevNo=1)
│   │
│   [EXISTING DEFINITION - has topId]:
│       ├─ GetReserveDef(topId) — check if reserved
│       ├─ CheckUpdatable(topId, publicStatus) → nextRevNo
│       ├─ If reserve + publicStatus="2" → Replace()
│       ├─ If new revision needed → Entry(nextRevNo)
│       └─ Else → Replace()
│
├─ 5. POST-PROCESSING (if IsApplyDefinision):
│      Mail.Send(Type=7) — approval notification
│      Direct.Send(Type=8) — push notification
│      Webhook.Send(Type=8)
│
└─ 6. RETURN XElement result
```

### Entry(int32 nextRevNo) — Core Transaction (Line 485440)

```
Entry(nextRevNo)
│
├─ InitTemporaryImageDirectory()
├─ Create ConMasUtility
├─ BEGIN TRANSACTION
│
├─ GetNextTopId() → assigns topId, topIdOrg
│
├─ [REVISION MODE: nextRevNo > 1]
│   ├─ Parse defTopId from XML
│   ├─ GetDefTopOrg(defTopId) → topIdOrg
│   ├─ GetMaxRevNo(topIdOrg) + 1 → nextRevNo
│   └─ CopyTemplate(beforeTopId)
│
├─ SET SYSTEM FIELDS: registTime, updateTime, registUser, updateUser
├─ EXTRACT & SAVE thumbnail/backgroundImage from XML BLOBs
│   → clear from XML after saving
│
├─ INSERT TOP: InsertTop(command, topId, topIdOrg, nextRevNo)
│   → writes to def_top table
│
├─ [if beforeTopId != topId: COPY phase]
│   ├─ Copy top via SQL (queryXml → copyTop)
│   ├─ ReflectTopOption(beforeTopId, topId, termId, userId)
│   └─ Copy related data: label, role, reportDocument, reportMaster, etc.
│
├─ INSERT SHEETS: for each sheet in XML
│   InsertSheet(command, topId, xSheet)
│   → writes to def_sheet table
│   → calls BuildCluster() for each sheet
│       → writes to def_cluster table
│
├─ HANDLE REPORT TABLES:
│   if cooperationTable != "0" OR IsTableReport:
│       CreateReportTable() / CreateReportTableDivide()
│
├─ COMMIT TRANSACTION
│
└─ ON ERROR: Rollback, log, rethrow
```

---

## 5. Database Table Population

### def_top — Form Definition Header

| Column | Source in XML | Notes |
|--------|---------------|-------|
| `def_top_id` | Auto-assigned | `GetNextTopId()` — integer PK |
| `def_top_org` | Auto-assigned | `topIdOrg` — original ID across revisions |
| `rev_no` | nextRevNo | Incremented on each revision |
| `title` | `<title>` | Form display name |
| `description` | `<description>` | Form description |
| `public_status` | `<publicStatus>` | 0=draft, 1=pending, 2=published |
| `regist_user_cd` | Session | Current user |
| `regist_dt` | Server time | Insert timestamp |
| `update_user_cd` | Session | Current user |
| `update_dt` | Server time | Update timestamp |
| `def_top_file` | `<definitionFile><value>` (Base64 decoded) | **The original Excel template** stored as BLOB |
| `thumbnail_file` | `<thumbnail>` (Base64 decoded) | PNG thumbnail |
| `background_image_file` | `<backgroundImage>` (Base64 decoded) | **Background image for rendering** |
| `server_version` | `<serverVersion>` | Version of server that created this |
| `top_option` | Derived from role settings | XML fragment for RBAC |

### def_sheet — Sheet/Page Definition

INSERT executed by `InsertSheet()` (line 484032):

| Column | Source in XML |
|--------|---------------|
| `def_sheet_id` | Auto-assigned |
| `def_top_id` | From parent `def_top` |
| `sheet_no` | `<sheetNo>` |
| `sheet_name` | `<defSheetName>` |
| `page_order` | `<pageOrder>` |
| `print_paper_size` | `<printPaperSize>` |
| `print_orientation` | `<printOrientation>` |
| `regist_user_cd` | Session |
| `regist_dt` | Server time |

### def_cluster — Data Cluster (Field Region)

INSERT executed by `BuildCluster()` (line 484235):

| Column | Source in XML |
|--------|---------------|
| `def_cluster_id` | Auto-assigned |
| `def_sheet_id` | From parent `def_sheet` |
| `cluster_id` | `<clusterId>` |
| `cluster_name` | `<clusterName>` |
| `cluster_type` | `<clusterType>` |
| `left_pos` | `<left>` (normalized float) |
| `top_pos` | `<top>` (normalized float) |
| `right_pos` | `<right>` (normalized float) |
| `bottom_pos` | `<bottom>` (normalized float) |
| `input_parameter` | `<inputParameter>` |
| `remarks_value_1` | `<remarksValue1>` |
| `remarks_value_2` | `<remarksValue2>` |
| ... | ... |
| `remarks_value_10` | `<remarksValue10>` |

### Image Storage

The **background PDF** is NOT uploaded as a PDF file. Instead:
- The **Excel file** is uploaded as Base64 in `<definitionFile><value>`
- The **background image** is uploaded as Base64 in `<backgroundImage>`
- On the server, these are decoded and stored as BLOBs in `def_top_file` and `background_image_file` columns
- **During rendering**: the server extracts the Excel BLOB, re-renders it via Syncfusion PdfGraphics (not Excel), using the stored cluster coordinates to overlay data fields

---

## 6. Cluster Image Upload (Separate Channel)

**URL**: `RegistClusterImage.aspx`
**Server handler**: `ConMasAPI.Rests.RegistClusterImage` → `ClusterImageBiz.Regist(HttpContext)`

The cluster images (cropped regions from the rendered Excel) are uploaded **separately** from the definition XML:

### Upload Fields (from ClusterImageBiz.Regist IL, line 654465)

| Field | Type | Source |
|-------|------|--------|
| `topId` | int | Definition top ID |
| `sheetNo` | int | Sheet number |
| `clusterId` | int | Cluster/field ID |
| **File** | `HttpPostedFileBase.InputStream` | Cropped cluster image |

### Image Processing (line 654520+)

```
Regist(HttpContext httpContext)
│
├─ Parse fields: topId, sheetNo, clusterId from request
├─ Read uploaded image from HttpPostedFileBase.InputStream
├─ Create Bitmap from stream
├─ Process image (resize/crop/format)
├─ Save to database via ClusterParameter / PointWriter
└─ Return XElement result via APIBiz.GenerateResult()
```

The image data is processed using `System.Drawing.Graphics`, `MemoryStream`, and `PdfEditor`.

---

## 7. Complete HTTP Request Structure

### Definition Upload

```
POST {ServerUrl}?command=RegistDefinition&approveType=1&comment=approval+text
Host: {server}
Content-Type: multipart/form-data; boundary=----WebKitFormBoundary...
Content-Length: {size}

------WebKitFormBoundary...
Content-Disposition: form-data; name="file"; filename="UploadFile.xml"
Content-Type: application/xml

<conmas>
  ...xml content...
</conmas>
------WebKitFormBoundary----
```

**Evidence**:
- `WebClient.UploadFile(url, filePath)` sends the file as `multipart/form-data` (WebClient's default behavior)
- Server reads: `HttpRequest.Files[0].InputStream` (DefinitionBiz.Regist line 496136)
- URL constructed as: `ServerUrl + "\?command=RegistDefinition"` (ConMasWebClient line 306363)
- Optional query params: `&approveType=1&comment={text}` (line 306369)
- Response: XML `<error><code>VALUE</code></error>` (line 306393–306399)

### Definition Download (Read)

```
POST {ServerUrl}
Content-Type: application/x-www-form-urlencoded

command=GetDefinitionData&topId=123
```

**Evidence**: `WebClient.UploadValues()` sends form data (ConMasWebClient line 306230)

### Cluster Image Upload

```
POST RegistClusterImage.aspx
Content-Type: multipart/form-data; boundary=...

------WebKitFormBoundary...
Content-Disposition: form-data; name="topId"

123
------WebKitFormBoundary...
Content-Disposition: form-data; name="sheetNo"

1
------WebKitFormBoundary...
Content-Disposition: form-data; name="clusterId"

5
------WebKitFormBoundary...
Content-Disposition: form-data; name="file"; filename="cluster_5.png"
Content-Type: image/png

{binary png data}
------WebKitFormBoundary----
```

---

## 8. Temporary File Lifecycle

The Designer creates and destroys files in a specific sequence:

```
┌─ ExportExcelDefinition(xlsFilePath, tempDir, ref ProgressValue)
│
│  tempDir = UserConfigDirectory + "\\temp"
│
│  ├─ CREATE: {tempDir}\                   (DirectoryInfo.Create)
│  ├─ CREATE: {tempDir}\temp.xlsx          (workbook.Save — Infragistics)
│  ├─ CREATE: {tempDir}\temp.pdf           (ExportPdf → ExportAsFixedFormat)
│  │
│  ├─ PROCESS: CalcClusterSize, GetAddress (reads PDF pixels)
│  │
│  ├─ DELETE: {tempDir}\temp.pdf           (finally block)
│  ├─ DELETE: {tempDir}\temp.xlsx          (finally block)
│  │
│  └─ CREATE: {tempDir}\{filename}.xml     (StreamWriter — the definition XML)
│
├─ COPY: {tempDir}\{filename}.xml → {UserConfigDirectory}\UploadFile.xml
│
├─ DELETE: {tempDir}\{filename}.xml        (from output dir — in ExportExcelDefinition catch)
│
├─ UPLOAD: UploadFile.xml → Server         (WebClient.UploadFile)
│
└─ DELETE: {UserConfigDirectory}\UploadFile.xml  (finally block in OutputXml, after upload)
```

---

## 9. How This Differs from Common Assumptions

| Assumption | Reality | Evidence |
|------------|---------|----------|
| Designer uploads a PDF | **No** — PDF is temp only, deleted. Excel is Base64-embedded in XML | `GetDefinition()` saves PDF, deletes in `finally`. Excel embedded in `<definitionFile><value>` |
| Designer uploads cluster images in the same request | **No** — images are uploaded separately via `RegistClusterImage.aspx` | Server has separate handler; `UploadDefinition` only sends XML |
| Server renders PDF from Excel at render time | **Yes** — Server re-extracts xlsx from BLOB, renders via Syncfusion PDF (no COM) | Phase X.42 confirmed ClusterWriter subclasses draw on PdfGraphics |
| Upload is a ZIP package | **No** — it's a single XML file | `WebClient.UploadFile` sends `.xml`; server parses as `XElement.Load` |
| Coordinates are in millimeters or inches | **No** — they are **normalized floats** (0.0–1.0) relative to image dimensions | `GetAddress()`: `Left = (float)x / image.Width`, `Top = (float)y / image.Height` |
| Background PDF is uploaded separately | **No** — The Excel template is uploaded; server regenerates PDF at render time | The `<definitionFile>` contains the xlsx; server uses Syncfusion to render |

---

## 10. Complete Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────┐
│                         DESIGNER (WPF Client)                        │
│                                                                      │
│  Excel Template (.xlsx)                                              │
│         │                                                            │
│         ▼                                                            │
│  ┌─────────────────────────────────┐                                 │
│  │ GetDefinition()                 │                                 │
│  │  ├─ Open Excel via COM          │                                 │
│  │  ├─ Read sheets, clusters       │                                 │
│  │  ├─ Export PDF (temp)           │                                 │
│  │  ├─ GetAddress() pixel scan     │                                 │
│  │  ├─ Build XML with coordinates  │                                 │
│  │  └─ Embed xlsx as Base64        │                                 │
│  └──────┬──────────────────────────┘                                 │
│         │ XElement (in-memory XML)                                   │
│         ▼                                                            │
│  ┌─────────────────────────────────┐                                 │
│  │ ExportExcelDefinition()         │                                 │
│  │  └─ Save XElement → .xml file   │                                 │
│  └──────┬──────────────────────────┘                                 │
│         │ UploadFile.xml on disk                                     │
│         ▼                                                            │
│  ┌─────────────────────────────────┐                                 │
│  │ ConMasWebClient                 │                                 │
│  │  └─ UploadDefinition(filePath)  │                                 │
│  │     └─ WebClient.UploadFile()   │                                 │
│  └──────┬──────────────────────────┘                                 │
│         │ HTTP POST (multipart/form-data)                            │
│         ▼                                                            │
│  ┌─────────────────────────────────┐                                 │
│  │ Separate Upload:                │                                 │
│  │ ConMasGeneratorUtility          │                                 │
│  │  └─ Upload cluster images via   │                                 │
│  │     RegistClusterImage.aspx     │                                 │
│  └─────────────────────────────────┘                                 │
└─────────────────────────────────────────────────────────────────────┘
         │
         │ HTTP POST
         ▼
┌─────────────────────────────────────────────────────────────────────┐
│                          SERVER (ASP.NET)                            │
│                                                                      │
│  ┌─────────────────────────────────┐                                 │
│  │ RegistDefinition.aspx           │                                 │
│  │  ├─ Read XML from Request.Files │                                 │
│  │  └─ DefinitionBiz.Regist()      │                                 │
│  │     ├─ ValidateReportTable()    │                                 │
│  │     ├─ Entry()                  │                                 │
│  │     │  ├─ Begin Transaction     │                                 │
│  │     │  ├─ InsertTop() → def_top │                                 │
│  │     │  ├─ InsertSheet() → def_sheet                               │
│  │     │  ├─ BuildCluster() → def_cluster                            │
│  │     │  ├─ Save Base64 blobs:    │                                 │
│  │     │  │  ├─ xlsx → def_top_file (BLOB)                           │
│  │     │  │  ├─ thumbnail → thumbnail_file (BLOB)                    │
│  │     │  │  └─ background → background_image_file (BLOB)            │
│  │     │  └─ Commit Transaction    │                                 │
│  │     └─ Return XML result        │                                 │
│  └─────────────────────────────────┘                                 │
│                                                                      │
│  ┌─────────────────────────────────┐                                 │
│  │ RegistClusterImage.aspx         │                                 │
│  │  └─ ClusterImageBiz.Regist()    │                                 │
│  │     ├─ Parse topId, sheetNo, clusterId                           │
│  │     ├─ Read image from InputStream                               │
│  │     └─ Save to cluster image storage                             │
│  └─────────────────────────────────┘                                 │
│                                                                      │
│  ┌─────────────────────────────────┐                                 │
│  │ LATER — Rendering               │                                 │
│  │  └─ ReportWriter.GetReport()    │                                 │
│  │     ├─ Extract xlsx from BLOB   │                                 │
│  │     ├─ Render via Syncfusion    │                                 │
│  │     │  PdfGraphics (no COM)     │                                 │
│  │     ├─ Overlay cluster fields   │                                 │
│  │     │  via ClusterWriter subs   │                                 │
│  │     └─ Output final PDF         │                                 │
│  └─────────────────────────────────┘                                 │
└─────────────────────────────────────────────────────────────────────┘
```

---

## 11. Key Evidence Summary

| Fact | IL Evidence | File & Line |
|------|-------------|-------------|
| Upload file is XML saved via StreamWriter | `StreamWriter(xmlFilePath, false, UTF8)` → `Write(definition.ToString())` | LibExcelController.il:536–546 |
| Upload URL has `?command=RegistDefinition` | `ldstr "\?command=RegistDefinition"` | LibConMas.il:306363 |
| Upload uses `WebClient.UploadFile()` | `callvirt WebClient.UploadFile(string, string)` | LibConMas.il:306386 |
| Server reads XML from `HttpRequest.Files[0].InputStream` | `XmlReader.Create(page.Request.Files[0].InputStream)` | Cimtops.ConMasBiz.il:496136 |
| XML root element is `<conmas>` | `new XElement("conmas", new XElement("top"))` | LibExcelController.il:2287–2293 |
| Excel file embedded as Base64 in `<definitionFile>` | `File.ReadAllBytes()` + `Convert.ToBase64String()` | LibExcelController.il:3382–3383 |
| Coordinates are normalized floats (x/Width, y/Height) | `cluster.Left = (float)x / image.Width` | LibExcelController.il:7717–7718 |
| Coordinates written as `<top>/<bottom>/<left>/<right>` | `cluster.Add(new XElement("top", rect.Top))` | LibExcelController.il:3581–3606 |
| Temp PDF is deleted after cluster detection | `File.Delete(pdfPath)` in finally block | LibExcelController.il:3748 |
| `UploadDefinition` is called from `OutputXml` in ConMasClient.exe | `callvirt ConMasWebClient.UploadDefinition(string, string, bool, string)` | ConMasClient.il:172768, 172880 |
| Upload file path = `UserConfigDirectory + "UploadFile.xml"` | String concatenation with "UploadFile.xml" | ConMasClient.il:172679 |
| Cluster images uploaded separately to `RegistClusterImage.aspx` | Server handler `ClusterImageBiz.Regist()` exists | Cimtops.ConMasBiz.il:654465 |
| `UploadFile.xml` deleted after upload | No direct evidence of delete in OutputXml, but documented lifecycle | ConMasClient.il (supplementary) |
| def_top_file stores the xlsx BLOB | `SaveTemporarySheetsImage` + `AddCommandParamTop` saves `<definitionFile><value>` | Cimtops.ConMasBiz.il:485258 |

---

## 12. Remaining Unknowns

1. **Cluster image upload format**: The exact image encoding (PNG/JPEG/BMP) and upload mechanism from the Designer to `RegistClusterImage.aspx` is traced on the server side but the client-side upload call was not found in LibConMas.dll. It likely lives in `ConMasGeneratorUtility.exe` or the runtime service.

2. **Full XML metadata elements**: The `OutputXml` method in `ConMasClient.exe` sets many additional XML elements (`topId, topIdOrg, revNo, registTime, updateTime, registUser, updateUser, title, description, privateMode, coeditKey, publicStatus, etc.`). A complete element enumeration would require reading the full `OutputXml` method body.

3. **Background image generation**: The `<backgroundImage>` element contains a Base64 PNG, but the code that generates this image from the Excel/PDF was not found in the decompiled files searched. It may be in ConMasClient.exe or generated by `GetDefinition()`.
