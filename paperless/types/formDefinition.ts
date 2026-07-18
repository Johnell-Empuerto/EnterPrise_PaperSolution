/**
 * TypeScript type definitions mirroring the C# FormDefinition model.
 * Used for round-trip Upload Excel / Export Excel workflow.
 */

export interface FormDefinition {
  workbook: WorkbookMetadata;
  sheets: SheetDefinition[];
  clusters: ClusterDefinition[];
  images: ImageDefinition[];
  metadata: Record<string, string>;
}

export interface WorkbookMetadata {
  title: string;
  author: string;
  created: string;
  modified: string;
  version: string;
  description: string;
}

export interface SheetDefinition {
  id: string;
  name: string;
  index: number;
  pageSettings: PageSettings;
  printArea: PrintAreaInfo | null;
  backgroundImage: string | null;
  backgroundWidth: number;
  backgroundHeight: number;
  thumbnail: string | null;
  rowHeights: Record<number, number>;
  columnWidths: Record<number, number>;
  mergedCells: MergedCellInfo[];
  freezePane: string | null;
  cellStyles: Record<string, CellStyleInfo>;
  cellValues: Record<string, string>;
  metadata: Record<string, string>;
}

export interface PageSettings {
  paperSize: string;
  orientation: string;
  widthPt: number;
  heightPt: number;
  leftMargin: number;
  topMargin: number;
  rightMargin: number;
  bottomMargin: number;
  centerHorizontally: boolean;
  centerVertically: boolean;
  zoom: number;
  fitToPagesWide: number;
  fitToPagesTall: number;
}

export interface PrintAreaInfo {
  address: string;
  leftPt: number;
  topPt: number;
  widthPt: number;
  heightPt: number;
  cols: number;
  rows: number;
}

export interface MergedCellInfo {
  address: string;
  cellAddress: string;
  leftPt: number;
  topPt: number;
  widthPt: number;
  heightPt: number;
}

export interface CellStyleInfo {
  fontName: string | null;
  fontSize: number | null;
  bold: boolean | null;
  italic: boolean | null;
  underline: boolean | null;
  color: string | null;
  fillColor: string | null;
  borderTop: string | null;
  borderBottom: string | null;
  borderLeft: string | null;
  borderRight: string | null;
  horizontalAlignment: string | null;
  verticalAlignment: string | null;
  wrapText: boolean | null;
}

export interface ClusterDefinition {
  clusterId: string;
  name: string;
  type: string;
  sheetId: string;
  cellAddress: string;
  left: number;
  right: number;
  top: number;
  bottom: number;
  leftPt: number;
  topPt: number;
  widthPt: number;
  heightPt: number;
  inputParameters: Record<string, string>;
  visibility: string;
  readonly: boolean;
  remarks: string;
  functions: string[];
  metadata: Record<string, string>;
}

export interface ImageDefinition {
  id: string;
  sheetId: string;
  name: string;
  leftPt: number;
  topPt: number;
  widthPt: number;
  heightPt: number;
  data: string | null;
  format: string;
}

// ── API Response Types ──

export interface UploadExcelResponse {
  success: boolean;
  message: string;
  data: {
    formDefinition: FormDefinition;
    fieldCount: number;
    sheetCount: number;
    templateId: string;
    workbookDownloadUrl: string;
  };
}

export interface OutputExcelResponse {
  success: boolean;
  message: string;
  data: {
    workbookPath: string;
    downloadUrl: string;
    fieldCount: number;
    sheetCount: number;
  };
}
