# Designer Loading Pipeline — Database → Screen

## Complete Load Sequence

### Step 1: Database Query (via ASP.NET WebForms API)
The Designer or Tablet loads a form by calling API endpoints (`.aspx` pages):

```
GET DefinitionDetailBiz.aspx?defTopId={id}
     ↓
1. Query def_top (form header)
   SELECT * FROM def_top WHERE def_top_id = {id}
   
2. Query def_label (labels/remarks)
   SELECT * FROM def_label WHERE def_top_id = {id}
   
3. Query def_cluster (clusters/fields)
   SELECT * FROM def_cluster WHERE def_top_id = {id}
   ORDER BY cluster_index, row, col
   
4. Query permissions (terminal/group access)
   
5. Query xml_data (full form XML)
   SELECT xml_data FROM def_top WHERE def_top_id = {id}

6. Query def_reference (reference data)
   SELECT * FROM def_reference WHERE def_top_id = {id}

7. Query def_cluster_role (cluster permissions)
   SELECT * FROM def_cluster_role WHERE def_top_id = {id}

8. Load background_image_file (PDF)
   SELECT background_image_file FROM def_top WHERE def_top_id = {id}
```

This loading order is confirmed from the SQL queries in:
- `DefinitionDetailBiz.xml` (tablet API)
- `DefinitionBiz.xml` (full definition)
- `ReportDetailBiz.xml` (report loading)

### Step 2: XML Parsing
The `xml_data` column contains the full form definition XML (conmas schema):

```xml
<conmas>
  <top>
    <defTopId>546</defTopId>
    <defTopName>FormTest - Copy</defTopName>
    <updateDefinitionApp>ConMasDesigner</updateDefinitionApp>
    <designerVersion>...</designerVersion>
    <sheets>
      <sheet>
        <width>612</width>
        <height>792</height>
        <clusters>
          <cluster>
            <sheetNo>1</sheetNo>
            <clusterId>0</clusterId>
            <name>samples</name>
            <type>KeyboardText</type>
            <left>0.3364706</left>
            <top>0.3845454</top>
            <right>0.4982353</right>
            <bottom>0.4218182</bottom>
            ...
          </cluster>
        </clusters>
      </sheet>
    </sheets>
  </top>
</conmas>
```

### Step 3: Background Rendering
```
PDF bytes (from background_image_file)
    │
    ├──► Temp file on disk
    ├──► ConvertPdfToImage() → System.Drawing.Image
    │     (DPI controlled by DpiPdfToImage / DpiPdfToBitmap config)
    │
    └──► WPF BitmapImage → Image control on canvas
```

Using one of:
- `O2S.Components.PDFRender4NET.WPF.dll` (tablet)
- `AHPDFLib10.dll` (Designer)
- `PdfiumViewer.dll` (Designer)
- `DevExpress.Pdf.*.dll` (Designer)

### Step 4: Cluster Object Creation

For each def_cluster record, the Designer creates a field object:

```
foreach (var cluster in defClusters)
{
    // Convert ratio coordinates to canvas pixels
    left_px   = cluster.left_position   * canvasWidth;
    top_px    = cluster.top_position    * canvasHeight;
    right_px  = cluster.right_position  * canvasWidth;
    bottom_px = cluster.bottom_position * canvasHeight;
    
    // Create appropriate control based on cluster type
    var control = CreateClusterControl(cluster.type);
    control.SetBounds(left_px, top_px, 
                      right_px - left_px, 
                      bottom_px - top_px);
    
    // Apply cluster metadata
    control.Name = cluster.name;
    control.ReadOnly = cluster.readOnly;
    control.InputParameters = ParseInputParameters(cluster.input_parameter);
    
    // Add to canvas
    canvas.Children.Add(control);
}
```

### Step 5: Yellow Overlay Creation

The yellow overlay rectangles are NOT the controls themselves — they are separate `CanvasChild+Rect` elements drawn behind or alongside the controls:

```
foreach (var cluster in defClusters)
{
    var overlay = new CanvasChild();
    overlay.Type = CanvasChildType.Rect;
    overlay.Bounds = new Rect(left_px, top_px, 
                              width_px, height_px);
    overlay.Fill = semiTransparentYellow;
    overlay.Stroke = darkYellow;
    overlay.StrokeThickness = 1;
    
    // Z-order behind controls
    Canvas.SetZIndex(overlay, 1);
    canvas.Children.Add(overlay);
}
```

### Step 6: User Interaction Layer

The Designer adds:
- Selection handling (mouse events on canvas + xamDataGrid)
- Resize handles (draggable at corners/edges)
- Grid snapping (alignment to nearest grid point)
- Context menus (right-click on clusters)
- Property editing (right panel with cluster properties)

## Key Objects Created in Memory

| Object | Source | Description |
|--------|--------|-------------|
| `FormDefinition` | def_top | Form header with all metadata fields |
| `SheetDefinition` | def_sheet | Per-sheet dimensions and metadata |
| `ClusterDefinition` | def_cluster | Per-field with coords, type, properties |
| `BackgroundImage` | background_image_file | PDF bytes → rendered bitmap |
| `FormXml` | xml_data | Full XML document (authoritative structure) |
| `ClusterParameters` | input_parameter | Type-specific parameter objects |
| `CanvasChild[]` | — | WPF visual elements on canvas |

## Database Entity Types (from LibConMas.ImageDb)

The `LibConMas.ImageDb` namespace defines typed parameter classes for each cluster type:

| Cluster Type | Parameter Class | Key Fields |
|-------------|-----------------|------------|
| Text | `TextClusterParameter` | InputRestriction, KeyboardType, TextClusterMode |
| Numeric | `NumericClusterParameter` | MinValue, MaxValue, DecimalPlaces |
| Check | `CheckClusterParameter` | CheckAllowMinCluster, CheckAllowMaxCluster |
| Select | `SelectClusterParameter` | SelectItems, MultiSelect |
| CalendarDate | `CalendarDateClusterParameter` | DateFormat, MinDate, MaxDate |
| QRCode | `QRCodeClusterParameter` | QRCodeFrom, QRCodeTo, QRCodeMessage |
| CodeReader | `CodeReaderClusterParameter` | BarcodeType, CodeReaderLines |
| DrawingImage | `DrawingImageClusterParameter` | DrawingPinNo, ImageCustomSize |
| FreeDraw | `FreeDrawClusterParameter` | BackgroundColor, BorderDetect |
| Handwriting | `HandwritingClusterParameter` | — |
| Image | `ImageClusterParameter` | ImageDpi, ImageCustomSize |
| Calculate | `CalculateClusterParameter` | CalculateAlignValue, CalculateColorValue |
| Action | (`Action*` fields) | ActionCommand, ActionURLValue, ActionPost |

## XML → Database Relationship

The `xml_data` and `def_cluster` serve different purposes:

| Aspect | xml_data | def_cluster |
|--------|----------|-------------|
| Coordinates | Yes (left/top/right/bottom) | Yes (left_position/right_position/top_position/bottom_position) |
| Names | Yes (cluster name) | Yes (name column) |
| Types | Yes (type string) | Yes (cluster_type column) |
| Input parameters | Yes (inputParameters) | Yes (input_parameter column) |
| Cell address | No | Yes (cell_addr column) |
| Sorting | Sequential in XML | ClusterIndex + Row + Col |
| Authoritative for | Form structure | Runtime coordinates + DB operations |
| Used by | Tablet form rendering | Tablet coordinate positioning |
