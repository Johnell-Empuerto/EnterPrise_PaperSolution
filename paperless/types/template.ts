export interface TemplatePageSetup {
  paperWidthPt: number;
  paperHeightPt: number;
  marginLeftIn: number;
  marginRightIn: number;
  marginTopIn: number;
  marginBottomIn: number;
  marginHeaderIn: number;
  marginFooterIn: number;
  centerHorizontally: boolean;
  centerVertically: boolean;
  orientation: "Portrait" | "Landscape";
  paperSize: number;
}

export interface TemplatePrintArea {
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
}

export interface TemplateMergedCell {
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
}

export interface TemplateCellStyle {
  fontName?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontColor?: string;
  fillColor?: string;
  borderTop?: string;
  borderBottom?: string;
  borderLeft?: string;
  borderRight?: string;
  horizontalAlignment?: "left" | "center" | "right";
  verticalAlignment?: "top" | "middle" | "bottom";
  wrapText?: boolean;
  indent?: number;
}

export interface TemplateComment {
  address: string;
  text: string;
}

export interface TemplateImage {
  anchorCol: number;
  anchorRow: number;
  offsetXPt?: number;
  offsetYPt?: number;
  widthPt: number;
  heightPt: number;
  imageData?: string; // base64 data URL
  imageUrl?: string;  // fallback URL if no data
  description?: string;
}

export interface TemplateModel {
  sheetName: string;
  pageSetup: TemplatePageSetup;
  printArea: TemplatePrintArea;
  columnWidths: number[];
  rowHeights: number[];
  hiddenColumns?: boolean[];
  hiddenRows?: boolean[];
  mergedCells: TemplateMergedCell[];
  cellValues: Record<string, string>;
  cellStyles: Record<string, TemplateCellStyle>;
  comments: TemplateComment[];
  images: TemplateImage[];
}
