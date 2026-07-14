"use client";

import { useMemo } from "react";
import type { TemplateModel } from "@/types/template";
import { ExcelCell } from "./ExcelCell";
import { ExcelMergedCell } from "./ExcelMergedCell";
import { ExcelImage } from "./ExcelImage";
import {
  getMergeSpan,
  isInsideMerge,
  getCellStyle,
  getCellValue,
  buildGridDimensions,
  cumulativeColWidth,
  cumulativeRowHeight,
  pt,
} from "./helpers";
import styles from "./ExcelRenderer.module.css";

export interface ExcelGridProps {
  template: TemplateModel;
  onCellClick?: (col: number, row: number) => void;
  onCellHover?: (col: number | null, row: number | null) => void;
}

export function ExcelGrid({ template, onCellClick, onCellHover }: ExcelGridProps) {
  const dims = useMemo(() => buildGridDimensions(template), [template]);

  const { printArea, mergedCells, cellStyles, cellValues, columnWidths, rowHeights, hiddenColumns, hiddenRows, images } = template;

  const gridTemplateColumns = dims.columnWidthsPt.map((w) => pt(w)).join(" ");
  const gridTemplateRows = dims.rowHeightsPt.map((h) => pt(h)).join(" ");

  // Build a mapping from logical (col,row) to grid column/row indices
  // accounting for hidden rows/cols that are skipped
  const visibleColIndex = useMemo(() => {
    const map = new Map<number, number>();
    let idx = 1;
    for (let c = printArea.startCol; c <= printArea.endCol; c++) {
      if (!hiddenColumns?.[c - 1]) {
        map.set(c, idx);
        idx++;
      }
    }
    return map;
  }, [printArea, hiddenColumns]);

  const visibleRowIndex = useMemo(() => {
    const map = new Map<number, number>();
    let idx = 1;
    for (let r = printArea.startRow; r <= printArea.endRow; r++) {
      if (!hiddenRows?.[r - 1]) {
        map.set(r, idx);
        idx++;
      }
    }
    return map;
  }, [printArea, hiddenRows]);

  const cells: React.ReactNode[] = [];

  for (let r = printArea.startRow; r <= printArea.endRow; r++) {
    // Skip hidden rows entirely
    if (hiddenRows?.[r - 1]) continue;

    for (let c = printArea.startCol; c <= printArea.endCol; c++) {
      // Skip hidden columns entirely
      if (hiddenColumns?.[c - 1]) continue;

      if (isInsideMerge(c, r, mergedCells)) continue;

      const span = getMergeSpan(c, r, mergedCells);

      const colIdx = visibleColIndex.get(c) ?? (c - printArea.startCol + 1);
      const rowIdx = visibleRowIndex.get(r) ?? (r - printArea.startRow + 1);

      const handleClick = onCellClick ? () => onCellClick(c, r) : undefined;
      const handleMouseEnter = onCellHover ? () => onCellHover(c, r) : undefined;
      const handleMouseLeave = onCellHover ? () => onCellHover(null, null) : undefined;

      if (span) {
        cells.push(
          <ExcelMergedCell
            key={`m${c}_${r}`}
            value={getCellValue(c, r, cellValues)}
            style={getCellStyle(c, r, cellStyles)}
            startCol={colIdx}
            startRow={rowIdx}
            spanCols={span.cols}
            spanRows={span.rows}
            onClick={handleClick}
            onMouseEnter={handleMouseEnter}
            onMouseLeave={handleMouseLeave}
          />
        );
      } else {
        const colWidth = dims.columnWidthsPt[colIdx - 1];
        const rowHeight = dims.rowHeightsPt[rowIdx - 1];
        cells.push(
          <ExcelCell
            key={`c${c}_${r}`}
            value={getCellValue(c, r, cellValues)}
            style={getCellStyle(c, r, cellStyles)}
            widthPt={colWidth}
            heightPt={rowHeight}
            gridColumn={colIdx}
            gridRow={rowIdx}
            onClick={handleClick}
            onMouseEnter={handleMouseEnter}
            onMouseLeave={handleMouseLeave}
          />
        );
      }
    }
  }

  // Render anchored images as absolutely positioned elements
  const imageElements = useMemo(() => {
    if (!images || images.length === 0) return null;
    return images.map((img, i) => {
      const leftPt =
        cumulativeColWidth(img.anchorCol, columnWidths, hiddenColumns) +
        (img.offsetXPt ?? 0);
      const topPt =
        cumulativeRowHeight(img.anchorRow, rowHeights, hiddenRows) +
        (img.offsetYPt ?? 0);
      return (
        <ExcelImage
          key={`img_${i}`}
          image={img}
          leftPt={leftPt}
          topPt={topPt}
        />
      );
    });
  }, [images, columnWidths, rowHeights, hiddenColumns, hiddenRows]);

  return (
    <div
      className={styles.grid}
      style={{
        gridTemplateColumns,
        gridTemplateRows,
        width: pt(dims.totalWidthPt),
      }}
    >
      {cells}
      {imageElements}
    </div>
  );
}
