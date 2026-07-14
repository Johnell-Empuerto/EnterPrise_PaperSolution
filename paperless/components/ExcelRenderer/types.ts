import type { TemplateCellStyle, TemplateMergedCell } from "@/types/template";

export interface GridCell {
  col: number;
  row: number;
  value: string;
  style: TemplateCellStyle;
  isMerged: boolean;
  mergeSpan?: { cols: number; rows: number };
}

export interface GridRow {
  rowIndex: number;
  heightPt: number;
  cells: GridCell[];
}

export interface CellAddress {
  col: number;
  row: number;
}

export interface GridDimensions {
  totalWidthPt: number;
  totalHeightPt: number;
  columnWidthsPt: number[];
  rowHeightsPt: number[];
}
