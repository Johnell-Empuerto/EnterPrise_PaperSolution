"use client";

import type { TemplateCellStyle } from "@/types/template";
import { parseExcelBorder } from "./helpers";
import styles from "./ExcelRenderer.module.css";

export interface ExcelCellProps {
  value: string;
  style: TemplateCellStyle;
  widthPt: number;
  heightPt: number;
  gridColumn: number;
  gridRow: number;
  onClick?: React.MouseEventHandler<HTMLDivElement>;
  onMouseEnter?: React.MouseEventHandler<HTMLDivElement>;
  onMouseLeave?: React.MouseEventHandler<HTMLDivElement>;
}

function borderStyle(
  side: "top" | "right" | "bottom" | "left",
  cellStyle: TemplateCellStyle
): string | undefined {
  const key = `border${side.charAt(0).toUpperCase() + side.slice(1)}` as keyof TemplateCellStyle;
  const raw = (cellStyle as Record<string, string | undefined>)[key];
  return parseExcelBorder(raw) ?? "0.5px solid #d0d0d0";
}

export function ExcelCell({
  value,
  style: cellStyle,
  widthPt,
  heightPt,
  gridColumn,
  gridRow,
  onClick,
  onMouseEnter,
  onMouseLeave,
}: ExcelCellProps) {
  const resolvedFontSize = cellStyle.fontSize ?? 11;
  const indentLevel = cellStyle.indent ?? 0;

  const inlineStyle: React.CSSProperties = {
    width: `${widthPt}pt`,
    height: `${heightPt}pt`,
    gridColumn,
    gridRow,
    fontFamily: cellStyle.fontName ?? "Calibri, sans-serif",
    fontSize: `${resolvedFontSize}pt`,
    fontWeight: cellStyle.bold ? 700 : 400,
    fontStyle: cellStyle.italic ? "italic" : "normal",
    textDecoration: cellStyle.underline ? "underline" : "none",
    color: cellStyle.fontColor ?? "#000000",
    backgroundColor: cellStyle.fillColor ?? "transparent",
    textAlign: (cellStyle.horizontalAlignment as React.CSSProperties["textAlign"]) ?? "left",
    verticalAlign: (cellStyle.verticalAlignment as React.CSSProperties["verticalAlign"]) ?? "middle",
    borderTop: borderStyle("top", cellStyle),
    borderRight: borderStyle("right", cellStyle),
    borderBottom: borderStyle("bottom", cellStyle),
    borderLeft: borderStyle("left", cellStyle),
    ...(cellStyle.wrapText
      ? { whiteSpace: "pre-wrap" as const, wordBreak: "break-word" as const }
      : {
          overflow: "hidden" as const,
          textOverflow: "ellipsis" as const,
          whiteSpace: "nowrap" as const,
        }),
    ...(indentLevel > 0
      ? { paddingLeft: `${indentLevel * resolvedFontSize * 1.8}pt` }
      : {}),
  };

  return (
    <div
      className={styles.cell}
      style={inlineStyle}
      data-col={gridColumn}
      data-row={gridRow}
      onClick={onClick}
      onMouseEnter={onMouseEnter}
      onMouseLeave={onMouseLeave}
      title={value || cellRefString(gridColumn, gridRow)}
    >
      {value}
    </div>
  );
}

function cellRefString(col: number, row: number): string {
  let letters = "";
  let c = col;
  while (c > 0) {
    c--;
    letters = String.fromCharCode(65 + (c % 26)) + letters;
    c = Math.floor(c / 26);
  }
  return `${letters}${row}`;
}
