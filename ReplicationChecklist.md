# Replication Checklist — Next.js Designer Requirements

This document enumerates every feature that must be replicated to build a Next.js Designer that behaves identically to the legacy ConMas Designer.

## Tier 1: Core Infrastructure (MUST HAVE)

### 1.1 Coordinate Calculation Engine
- [ ] **Column-width summation**: Sum column widths in points from OpenXML (not COM)
- [ ] **Row-height summation**: Sum row heights in points from OpenXML
- [ ] **OpenXML width conversion**: Convert `Column.Width` (character-based) to points
  - Note: OpenXML column widths use a different formula than COM `Range.Width`
  - Formula: `width_pt = (width_chars * charWidth_pt + padding_pt)` where `charWidth_pt` depends on NormalStyle font
- [ ] **MergeArea detection**: Find merged cell ranges
- [ ] **Comment reading**: Extract comments from Excel (line 0 = name, 1 = type, 2 = index)
- [ ] **Cluster bounds**: From merge area or single cell
- [ ] **Normalization**: Divide by page dimensions (PDF-derived, not PageSetup)
- [ ] **RoundEx**: Banker's rounding to 7 decimal places
- [ ] **Gap compensation**: Recover gapH/gapV from analysis

### 1.2 Page Setup Handling
- [ ] **Print Area**: Read and parse `$A$1:$D$12` format
- [ ] **Page size**: Letter (612×792), A4 (595×842), etc.
- [ ] **Margins**: Left, Right, Top, Bottom margins
- [ ] **Header/Footer**: Sizes and margins
- [ ] **CenterHorizontally**: Flag handling
- [ ] **CenterVertically**: Flag handling
- [ ] **Zoom/Scaling**: 100%, FitToPagesWide, FitToPagesTall
- [ ] **Hidden rows/columns**: Skip in summation
- [ ] **Page breaks**: Horizontal/vertical page break handling

### 1.3 PDF Background Generation
- [ ] **ExportAsFixedFormat**: Excel to PDF conversion
  - Option A: Server-side Excel COM (requires Office installation)
  - Option B: OpenXML + PDF library (e.g., jsPDF, PDFKit)
  - Option C: Excel JavaScript API (client-side, Excel Online)
- [ ] **PDF storage**: Store PDF bytes in database
- [ ] **PDF retrieval**: Load and serve PDF bytes from DB
- [ ] **PDF rendering**: Render PDF pages to canvas images
  - PDF.js (client-side browser rendering)
  - Or server-side PDF→PNG conversion

### 1.4 Database Interface
- [ ] **Read def_top**: Form metadata and xml_data
- [ ] **Read def_sheet**: Sheet dimensions
- [ ] **Read def_cluster**: Cluster definitions with coordinates
- [ ] **Read background_image_file**: PDF bytes
- [ ] **Write def_top**: INSERT for new publish
- [ ] **Write def_sheet**: Per-sheet INSERT
- [ ] **Write def_cluster**: Per-cluster INSERT
- [ ] **UPDATE xml_data**: Full form XML generation
- [ ] **UPDATE background_image_file**: PDF bytes
- [ ] **INSERT def_current**: Publish completion marker
- [ ] **Transaction**: All DB operations in single transaction

### 1.5 XML Generation
- [ ] **conmas root**: `<conmas>` with header
- [ ] **top section**: defTopId, name, version, type, metadata
- [ ] **sheet section**: Width, height, clusters
- [ ] **cluster section**: 
  - [ ] clusterId, sheetNo
  - [ ] name (from comment line 0)
  - [ ] type (from comment line 1)
  - [ ] left, top, right, bottom (normalized ratios)
  - [ ] cellAddress (merge range reference like `$A$1:$B$2`)
  - [ ] value (default value)
  - [ ] isHidden, isHiddenDesigner
  - [ ] readOnly
  - [ ] external
  - [ ] displayValue
  - [ ] cooperationCluster
  - [ ] pinNo, pinValue
  - [ ] inputParameters
  - [ ] remarks 1-10
  - [ ] cellAddress
- [ ] **Variable definitions**: nameParts, labels, networks, references
- [ ] **Grammar**: Cluster grammar definitions (voice input, select items)

## Tier 2: Designer Rendering (SHOULD HAVE)

### 2.1 Canvas
- [ ] **ZoomableCanvas**: Custom canvas component with scale/pan
- [ ] **Zoom**: Slider-controlled zoom (mouse wheel optional)
- [ ] **Pan**: Click-drag to pan
- [ ] **ScaleX/ScaleY**: Independent axis scaling
- [ ] **Coordinate transform**: Ratio → pixel conversion at any zoom level

### 2.2 Background Image
- [ ] **PDF rasterization**: Render PDF page to canvas image
  - Using PDF.js or server-side rendering
- [ ] **DPI control**: Render at configurable DPI (default 96/200/300)
- [ ] **Scale to fit**: Fit PDF page to canvas viewport
- [ ] **Multi-page**: Page navigation for multi-page forms

### 2.3 Cluster Overlays
- [ ] **Yellow rectangles**: Semi-transparent yellow overlay per cluster
- [ ] **Border**: Dark yellow/amber 1px border
- [ ] **Labels**: Cluster name label on overlay
- [ ] **Type indicator**: Show cluster type icon or text
- [ ] **Z-ordering**: Proper layering (background → overlays → selection → handles)

### 2.4 Selection
- [ ] **Click selection**: Click cluster to select
- [ ] **Multi-select**: Ctrl+click or shift+click for multiple
- [ ] **Selection color**: Different color for selected (e.g., light blue)
- [ ] **Selection list**: Side panel with cluster list
- [ ] **Keyboard navigation**: Arrow keys to move between clusters
- [ ] **Range selection**: Click-drag to select multiple clusters

### 2.5 Resize Handles
- [ ] **Corner handles**: 4 corner drag handles
- [ ] **Edge handles**: 4 edge midpoint drag handles
- [ ] **Snap to grid**: Position snapping during resize
- [ ] **Min/max size**: Constrain minimum/maximum cluster size

### 2.6 Grid Snapping
- [ ] **Grid display**: Optional grid lines
- [ ] **Snap threshold**: Configurable alignment threshold
- [ ] **Cell-aligned snap**: Snap to Excel cell boundaries
- [ ] **Device pixel snap**: Snap to device pixels for crisp rendering

### 2.7 Property Editing
- [ ] **Name**: Edit cluster name
- [ ] **Type**: Change cluster type (with validation)
- [ ] **Coordinates**: Manual coordinate input (ratios)
- [ ] **Read-only**: Toggle read-only flag
- [ ] **Hidden**: Toggle hidden flag
- [ ] **Input parameters**: Edit type-specific parameters
- [ ] **Remarks**: Edit 10 remark fields
- [ ] **Cluster index**: Sort order

## Tier 3: Advanced Features (NICE TO HAVE)

### 3.1 Cluster Types
- [ ] **All 38 cluster types** (from ClusterType enum):
  KeyboardText, Handwriting, FixedText, FreeText, FreeDraw, Numeric, InputNumeric, Calculate, TimeCalculate, NumberHours, Check, MultipleChoiceNumber, MCNCalculate, Select, MultiSelect, Date, CalendarDate, Time, Image, Registration, RegistrationDate, LatestUpdate, LatestUpdateDate, Create, Inspect, Approve, Gps, CodeReader, QRCode, SelectMaster, Action, LoginUser, DrawingImage, DrawingPinNo, PinItemTableNo, AudioRecording, Scandit, EdgeOCR

### 3.2 Table Support
- [ ] **Table definition**: Column headers, row headers
- [ ] **Column types**: Numeric, date, text, calculate
- [ ] **Table rendering**: Grid layout with clusters in cells
- [ ] **Table editing**: Add/remove columns and rows

### 3.3 Network Routing
- [ ] **Cluster networks**: Next-field navigation rules
- [ ] **Conditional routing**: Skip logic based on values
- [ ] **Auto-input**: Automatic field input on navigation

### 3.4 Multi-Language
- [ ] **UI language**: Japanese, English, Chinese, Traditional Chinese
- [ ] **Cluster type names**: Localized per language
- [ ] **Form labels**: Multi-language label support

### 3.5 EdgeOCR
- [ ] **OCR regions**: Define reading areas on background
- [ ] **Character type rules**: Regex, length, custom char sets
- [ ] **Date determination**: Date parsing rules
- [ ] **Data mapping**: OCR output to cluster mapping

### 3.6 Scandit (Barcode)
- [ ] **Barcode types**: All supported symbologies
- [ ] **Scan regions**: Overlay-based scan areas
- [ ] **Data decomposition**: GS1 application ID parsing
- [ ] **Output mapping**: Scan results to clusters

## Tier 4: Verification (MUST HAVE)

### 4.1 Regression Testing
- [ ] **Template 546**: 6 clusters — verify all coordinates match DB within 0.0001
- [ ] **Template 547**: 5 clusters — verify all coordinates match DB
- [ ] **Template 548**: 2 clusters — verify all coordinates match DB
- [ ] **Empty name handling**: Template 547 auto-name generation "Cluster0"-"Cluster4"
- [ ] **Cluster sort order**: ClusterIndex → Row → Col

### 4.2 Database Verification
- [ ] **Coord comparison**: Compare calculated ratios with def_cluster values
- [ ] **XML comparison**: Compare generated XML with stored xml_data
- [ ] **Background comparison**: Compare generated PDF with stored background (after PDF→raster)

## Architecture Decisions for Next.js Implementation

### Option A: Server-Side Excel Processing
```
Browser (Next.js)
    │
    ├──► Upload Excel file
    ├──► Server: OpenXML parsing (ExcelJS, ClosedXML, etc.)
    ├──► Server: Coordinate calculation
    ├──► Server: PDF generation (ExcelJS + PDF lib, or LibreOffice)
    ├──► Server: Database writes
    └──► Return form definition to browser
```

**Pros**: No client-side Excel dependency, works in all browsers
**Cons**: Complex server-side processing, PDF generation quality gap

### Option B: Client-Side Processing (Excel Online / Office JS)
```
Browser (Next.js with Office JS Add-In)
    │
    ├──► Office JS API reads workbook
    ├──► Client-side coordinate calculation
    ├──► Client-side preview rendering
    ├──► API calls to server for DB operations
    └──► Server only for persistence
```

**Pros**: Real Excel rendering, no PDF generation needed
**Cons**: Requires Excel Online/Desktop, Office JS API limitations

### Option C: Hybrid
```
Browser (Next.js) ←→ Server
    │
    ├──► Client: Upload workbook
    ├──► Server: Parse with Office COM or OpenXML library
    ├──► Client: Canvas-based preview
    ├──► Server: PDF generation via LibreOffice headless
    └──► Server: Database persistence
```

**Recommended approach** for maximum compatibility.

## Critical Path Items

1. **OpenXML width/height parsing** — Must convert Excel column widths from character-based to points using the exact formula
2. **Coordinate normalization** — Must determine exact page dimension source (PDF, PageSetup, or Sheet totals)
3. **GetClusterSize decompilation** — Highest priority unknown; must recover the exact algorithm
4. **PDF generation** — Must produce pixel-identical PDFs to pass verification
5. **PDF→image rendering** — Must render PDF at matching DPI for background comparison

## Testing Strategy

1. Parse all 3 templates with OpenXML, extract raw column/row measurements
2. Compare COM widths vs OpenXML widths for each template
3. Build coordinate calculation that produces matching def_cluster values
4. Generate matching XML structure
5. Generate matching PDF (can use same Excel COM ExportAsFixedFormat initially)
6. Verify round-trip: New publish → DB → Designer load → coordinate match
