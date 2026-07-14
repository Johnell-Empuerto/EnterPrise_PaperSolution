import type { TemplateModel, TemplateMergedCell, TemplateCellStyle } from "@/types/template";
import type { GridDimensions } from "./types";

export function pt(value: number): string {
  return `${value}pt`;
}

export function ptToPx(value: number): number {
  return value * (96 / 72);
}

export function parseCellRef(ref: string): { col: number; row: number } | null {
  const match = ref.toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  let col = 0;
  for (const ch of match[1]) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  return { col, row: parseInt(match[2], 10) };
}

export function cellRef(col: number, row: number): string {
  let letters = "";
  let c = col;
  while (c > 0) {
    c--;
    letters = String.fromCharCode(65 + (c % 26)) + letters;
    c = Math.floor(c / 26);
  }
  return `${letters}${row}`;
}

export function getMergeSpan(
  col: number,
  row: number,
  mergedCells: TemplateMergedCell[]
): { cols: number; rows: number } | null {
  for (const mc of mergedCells) {
    if (mc.startCol === col && mc.startRow === row) {
      return { cols: mc.endCol - mc.startCol + 1, rows: mc.endRow - mc.startRow + 1 };
    }
  }
  return null;
}

export function isInsideMerge(
  col: number,
  row: number,
  mergedCells: TemplateMergedCell[]
): boolean {
  for (const mc of mergedCells) {
    if (
      col >= mc.startCol &&
      col <= mc.endCol &&
      row >= mc.startRow &&
      row <= mc.endRow
    ) {
      return !(mc.startCol === col && mc.startRow === row);
    }
  }
  return false;
}

export function getCellStyle(
  col: number,
  row: number,
  cellStyles: Record<string, TemplateCellStyle>
): TemplateCellStyle {
  const ref = cellRef(col, row);
  return cellStyles[ref] ?? {};
}

export function getCellValue(
  col: number,
  row: number,
  cellValues: Record<string, string>
): string {
  const ref = cellRef(col, row);
  return cellValues[ref] ?? "";
}

/**
 * Parses an Excel border style string into a CSS border shorthand.
 *
 * Excel stores borders as strings like:
 *   "thin black"       → 1px solid #000000
 *   "medium #4472C4"   → 2px solid #4472C4
 *   "thick"            → 3px solid #000000
 *   "double"           → 3px double #000000
 *   "hair"             → 0.5px solid #000000
 *   "dotted"           → 1px dotted #000000
 *   "dashed"           → 1px dashed #000000
 *   "none" / ""        → undefined (no border)
 */
export function parseExcelBorder(borderStr: string | undefined): string | undefined {
  if (!borderStr || borderStr === "none" || borderStr === "") return undefined;

  const ExcelLineWidths: Record<string, string> = {
    thin: "1px",
    medium: "2px",
    thick: "3px",
    hair: "0.5px",
  };

  const ExcelLineStyles: Record<string, string> = {
    thin: "solid",
    medium: "solid",
    thick: "solid",
    double: "double",
    hair: "solid",
    dotted: "dotted",
    dashed: "dashed",
    dashdot: "dashed",
    dashdotdot: "dotted",
    mediumdashed: "dashed",
    mediumdashdot: "dashed",
    mediumdashdotdot: "dotted",
    slantdashdot: "dashed",
  };

  const lower = borderStr.toLowerCase().trim();

  // Extract style keyword (first word)
  const parts = lower.split(/\s+/);
  const styleWord = parts[0];

  const width = ExcelLineWidths[styleWord] ??
    (styleWord === "double" ? "3px" :
     styleWord.includes("medium") || styleWord.includes("thick") ? "2px" : "1px");

  const style = ExcelLineStyles[styleWord] ?? "solid";

  // Extract color (remaining parts after style, supports multi-word colors)
  let color = "#000000";
  if (parts.length > 1) {
    const colorStr = parts.slice(1).join(" ");
    if (colorStr.startsWith("#") || colorStr.startsWith("rgb")) {
      color = colorStr;
    } else {
      // Named color — map common ones
      const namedColors: Record<string, string> = {
        black: "#000000",
        white: "#ffffff",
        red: "#ff0000",
        green: "#00b050",
        blue: "#0070c0",
        yellow: "#ffff00",
        orange: "#ed7d31",
        purple: "#7030a0",
        gray: "#808080",
        "dark blue": "#002060",
        "dark green": "#006100",
        "dark red": "#8b0000",
        "light blue": "#8db4e2",
        "light green": "#70ad47",
        "light yellow": "#ffff99",
      };
      color = namedColors[colorStr.toLowerCase()] ?? colorStr;
    }
  }

  return `${width} ${style} ${color}`;
}

export function buildGridDimensions(template: TemplateModel): GridDimensions {
  const { printArea, columnWidths, rowHeights, hiddenColumns, hiddenRows } = template;

  const colSlice = columnWidths.slice(printArea.startCol - 1, printArea.endCol);
  const hiddenColSlice = hiddenColumns
    ? hiddenColumns.slice(printArea.startCol - 1, printArea.endCol)
    : undefined;

  const columnWidthsPt: number[] = [];
  let totalWidthPt = 0;
  for (let i = 0; i < colSlice.length; i++) {
    if (hiddenColSlice?.[i]) continue; // skip hidden columns
    columnWidthsPt.push(colSlice[i]);
    totalWidthPt += colSlice[i];
  }

  const rowHeightsPt: number[] = [];
  let totalHeightPt = 0;
  for (let r = printArea.startRow; r <= printArea.endRow; r++) {
    if (hiddenRows?.[r - 1]) continue; // skip hidden rows
    const h = rowHeights[r - 1] ?? 15;
    rowHeightsPt.push(h);
    totalHeightPt += h;
  }

  return { totalWidthPt, totalHeightPt, columnWidthsPt, rowHeightsPt };
}

/**
 * Calculates the print origin (top-left of content) in points.
 * Handles centering and margin-based positioning.
 */
export function originFromMargins(template: TemplateModel): { xPt: number; yPt: number } {
  const { pageSetup } = template;
  const dims = buildGridDimensions(template);
  const contentW = dims.totalWidthPt;
  const contentH = dims.totalHeightPt;

  const marginLeftPt = pageSetup.marginLeftIn * 72;
  const marginTopPt = pageSetup.marginTopIn * 72;
  const marginRightPt = pageSetup.marginRightIn * 72;
  const marginBottomPt = pageSetup.marginBottomIn * 72;

  const usableW = pageSetup.paperWidthPt - marginLeftPt - marginRightPt;
  const usableH = pageSetup.paperHeightPt - marginTopPt - marginBottomPt;

  return {
    xPt: pageSetup.centerHorizontally
      ? marginLeftPt + (usableW - contentW) / 2
      : marginLeftPt,
    yPt: pageSetup.centerVertically
      ? marginTopPt + (usableH - contentH) / 2
      : marginTopPt,
  };
}

/**
 * Calculates the cumulative column offset from the start of the grid
 * for a given column index (1-based).
 */
export function cumulativeColWidth(
  col: number,
  columnWidths: number[],
  hiddenColumns?: boolean[]
): number {
  let total = 0;
  for (let i = 0; i < col - 1; i++) {
    if (hiddenColumns?.[i]) continue;
    total += columnWidths[i] ?? 0;
  }
  return total;
}

/**
 * Calculates the cumulative row offset from the start of the grid
 * for a given row index (1-based).
 */
export function cumulativeRowHeight(
  row: number,
  rowHeights: number[],
  hiddenRows?: boolean[]
): number {
  let total = 0;
  for (let i = 0; i < row - 1; i++) {
    if (hiddenRows?.[i]) continue;
    total += rowHeights[i] ?? 15;
  }
  return total;
}


