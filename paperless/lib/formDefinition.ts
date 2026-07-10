export interface Size {
  width: number;
  height: number;
}

export interface Point {
  x: number;
  y: number;
}

export interface Margin {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface PageSettings {
  paperSize: string;
  orientation: "portrait" | "landscape";
  widthPt: number;
  heightPt: number;
  margins: Margin;
  centerHorizontally: boolean;
  centerVertically: boolean;
  zoom: number;
  fitToPagesWide: number;
  fitToPagesTall: number;
}

export interface PrintArea {
  address: string;
  leftPt: number;
  topPt: number;
  widthPt: number;
  heightPt: number;
  cols: number;
  rows: number;
}

export interface MergedCell {
  address: string;
  cellAddress: string;
  leftPt: number;
  topPt: number;
  widthPt: number;
  heightPt: number;
}

export interface CellStyle {
  fontName?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  fillColor?: string;
  borderTop?: string;
  borderBottom?: string;
  borderLeft?: string;
  borderRight?: string;
  horizontalAlignment?: string;
  verticalAlignment?: string;
  wrapText?: boolean;
}

export interface SheetDefinition {
  id: string;
  name: string;
  index: number;
  pageSettings: PageSettings;
  printArea: PrintArea | null;
  backgroundImage: string | null;
  backgroundWidth: number;
  backgroundHeight: number;
  thumbnail: string | null;
  rowHeights: Record<number, number>;
  columnWidths: Record<number, number>;
  mergedCells: MergedCell[];
  freezePane: string | null;
  cellStyles: Record<string, CellStyle>;
  cellValues: Record<string, string>;
  metadata: Record<string, string>;
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
  data: string;
  format: string;
}

export interface WorkbookMetadata {
  title: string;
  author: string;
  created: string;
  modified: string;
  version: string;
  description: string;
}

export interface FormDefinition {
  workbook: WorkbookMetadata;
  sheets: SheetDefinition[];
  clusters: ClusterDefinition[];
  images: ImageDefinition[];
  metadata: Record<string, string>;
}