"use client";

import type { TemplateModel, TemplateMergedCell, TemplateImage } from "@/types/template";
import { cellRef, cumulativeColWidth, cumulativeRowHeight } from "../ExcelRenderer/helpers";

export interface DebugOverlayProps {
  template: TemplateModel;
  showGridLines: boolean;
  showCellCoords: boolean;
  showMergeBounds: boolean;
  showImageBounds: boolean;
  showOrigin: boolean;
  showPrintArea: boolean;
  showMargins: boolean;
}

export function DebugOverlay({
  template,
  showGridLines,
  showCellCoords,
  showMergeBounds,
  showImageBounds,
  showOrigin,
  showPrintArea,
  showMargins,
}: DebugOverlayProps) {
  const { printArea, pageSetup, columnWidths, rowHeights, mergedCells, hiddenColumns, hiddenRows, images } = template;

  // Build cumulative column boundaries (in pt from grid start)
  const colBoundaries: number[] = [0];
  const visibleColBoundaries: number[] = [0];
  for (let c = printArea.startCol; c <= printArea.endCol; c++) {
    const w = columnWidths[c - 1] ?? 0;
    colBoundaries.push(colBoundaries[colBoundaries.length - 1] + w);
    if (!hiddenColumns?.[c - 1]) {
      visibleColBoundaries.push(visibleColBoundaries[visibleColBoundaries.length - 1] + w);
    }
  }

  // Build cumulative row boundaries (in pt from grid start)
  const rowBoundaries: number[] = [0];
  const visibleRowBoundaries: number[] = [0];
  for (let r = printArea.startRow; r <= printArea.endRow; r++) {
    const h = rowHeights[r - 1] ?? 15;
    rowBoundaries.push(rowBoundaries[rowBoundaries.length - 1] + h);
    if (!hiddenRows?.[r - 1]) {
      visibleRowBoundaries.push(visibleRowBoundaries[visibleRowBoundaries.length - 1] + h);
    }
  }

  const totalW = visibleColBoundaries[visibleColBoundaries.length - 1];
  const totalH = visibleRowBoundaries[visibleRowBoundaries.length - 1];

  const overlayStyle: React.CSSProperties = {
    position: "absolute",
    inset: 0,
    pointerEvents: "none",
    userSelect: "none",
    zIndex: 10,
  };

  const elements: React.ReactNode[] = [];

  // ── Grid Lines ──
  if (showGridLines) {
    // Vertical lines
    for (let i = 1; i < visibleColBoundaries.length - 1; i++) {
      const x = visibleColBoundaries[i];
      elements.push(
        <div
          key={`vg_${i}`}
          style={{
            position: "absolute",
            left: `${x}pt`,
            top: 0,
            width: "0.5px",
            height: `${totalH}pt`,
            background: "rgba(255, 0, 0, 0.2)",
          }}
        />
      );
    }
    // Horizontal lines
    for (let i = 1; i < visibleRowBoundaries.length - 1; i++) {
      const y = visibleRowBoundaries[i];
      elements.push(
        <div
          key={`hg_${i}`}
          style={{
            position: "absolute",
            left: 0,
            top: `${y}pt`,
            height: "0.5px",
            width: `${totalW}pt`,
            background: "rgba(255, 0, 0, 0.2)",
          }}
        />
      );
    }
  }

  // ── Cell Coordinates ──
  if (showCellCoords) {
    for (let r = printArea.startRow; r <= printArea.endRow; r++) {
      if (hiddenRows?.[r - 1]) continue;
      for (let c = printArea.startCol; c <= printArea.endCol; c++) {
        if (hiddenColumns?.[c - 1]) continue;
        const colOffset = cumulativeColWidth(c, columnWidths, hiddenColumns);
        const rowOffset = cumulativeRowHeight(r, rowHeights, hiddenRows);
        elements.push(
          <div
            key={`cc_${c}_${r}`}
            style={{
              position: "absolute",
              left: `${colOffset}pt`,
              top: `${rowOffset}pt`,
              fontSize: "7pt",
              fontFamily: "monospace",
              color: "rgba(0, 120, 200, 0.6)",
              lineHeight: 1,
              padding: "1px 2px",
              background: "rgba(255,255,255,0.5)",
              borderRadius: "1px",
            }}
          >
            {cellRef(c, r)}
          </div>
        );
      }
    }
  }

  // ── Merge Bounds ──
  if (showMergeBounds) {
    const mergeColors = [
      "rgba(255, 165, 0, 0.3)",
      "rgba(255, 100, 100, 0.3)",
      "rgba(100, 200, 255, 0.3)",
      "rgba(100, 255, 100, 0.3)",
    ];
    mergedCells.forEach((mc: TemplateMergedCell, i: number) => {
      const left = cumulativeColWidth(mc.startCol, columnWidths, hiddenColumns);
      const top = cumulativeRowHeight(mc.startRow, rowHeights, hiddenRows);
      const right = cumulativeColWidth(mc.endCol + 1, columnWidths, hiddenColumns);
      const bottom = cumulativeRowHeight(mc.endRow + 1, rowHeights, hiddenRows);
      const color = mergeColors[i % mergeColors.length];

      elements.push(
        <div
          key={`mb_${i}`}
          style={{
            position: "absolute",
            left: `${left}pt`,
            top: `${top}pt`,
            width: `${right - left}pt`,
            height: `${bottom - top}pt`,
            background: color,
            border: "1.5px solid rgba(255, 165, 0, 0.8)",
            borderRadius: "1px",
          }}
          title={`Merge: ${cellRef(mc.startCol, mc.startRow)}-${cellRef(mc.endCol, mc.endRow)}`}
        />
      );
    });
  }

  // ── Image Bounds ──
  if (showImageBounds && images) {
    (images as TemplateImage[]).forEach((img: TemplateImage, i: number) => {
      const left = cumulativeColWidth(img.anchorCol, columnWidths, hiddenColumns) + (img.offsetXPt ?? 0);
      const top = cumulativeRowHeight(img.anchorRow, rowHeights, hiddenRows) + (img.offsetYPt ?? 0);

      elements.push(
        <div
          key={`ib_${i}`}
          style={{
            position: "absolute",
            left: `${left}pt`,
            top: `${top}pt`,
            width: `${img.widthPt}pt`,
            height: `${img.heightPt}pt`,
            border: "1.5px solid rgba(255, 0, 255, 0.8)",
            background: "rgba(255, 0, 255, 0.1)",
            borderRadius: "1px",
          }}
          title={`Image ${i + 1}: ${img.widthPt}×${img.heightPt}pt at (${img.anchorCol},${img.anchorRow})`}
        />
      );
    });
  }

  // ── Origin Dot ──
  if (showOrigin) {
    const originX = cumulativeColWidth(printArea.startCol, columnWidths, hiddenColumns);
    const originY = cumulativeRowHeight(printArea.startRow, rowHeights, hiddenRows);

    elements.push(
      <div
        key="origin"
        style={{
          position: "absolute",
          left: `${originX}pt`,
          top: `${originY}pt`,
          width: "8pt",
          height: "8pt",
          borderRadius: "50%",
          background: "rgba(0, 100, 255, 0.9)",
          transform: "translate(-4pt, -4pt)",
          zIndex: 20,
        }}
        title={`Origin: (${originX.toFixed(1)}, ${originY.toFixed(1)}) pt`}
      />
    );
  }

  // ── Print Area Bounds ──
  if (showPrintArea) {
    const left = cumulativeColWidth(printArea.startCol, columnWidths, hiddenColumns);
    const top = cumulativeRowHeight(printArea.startRow, rowHeights, hiddenRows);
    const right = cumulativeColWidth(printArea.endCol + 1, columnWidths, hiddenColumns);
    const bottom = cumulativeRowHeight(printArea.endRow + 1, rowHeights, hiddenRows);

    elements.push(
      <div
        key="printArea"
        style={{
          position: "absolute",
          left: `${left}pt`,
          top: `${top}pt`,
          width: `${right - left}pt`,
          height: `${bottom - top}pt`,
          border: "1.5px dashed rgba(255, 0, 0, 0.7)",
          zIndex: 5,
        }}
        title={`Print Area: ${cellRef(printArea.startCol, printArea.startRow)}-${cellRef(printArea.endCol, printArea.endRow)}`}
      />
    );
  }

  // ── Margins ──
  if (showMargins) {
    const marginLeftPt = pageSetup.marginLeftIn * 72;
    const marginTopPt = pageSetup.marginTopIn * 72;
    const marginRightPt = pageSetup.marginRightIn * 72;
    const marginBottomPt = pageSetup.marginBottomIn * 72;
    const paperW = pageSetup.paperWidthPt;
    const paperH = pageSetup.paperHeightPt;

    // Compute print origin (where grid starts within the page)
    const dims = { totalWidthPt: totalW, totalHeightPt: totalH };
    const usableW = paperW - marginLeftPt - marginRightPt;
    const usableH = paperH - marginTopPt - marginBottomPt;
    const originX = pageSetup.centerHorizontally
      ? marginLeftPt + (usableW - dims.totalWidthPt) / 2
      : marginLeftPt;
    const originY = pageSetup.centerVertically
      ? marginTopPt + (usableH - dims.totalHeightPt) / 2
      : marginTopPt;

    // Draw margin lines offset from the page origin
    // (Since the overlay is relative to the grid, we show margins relative to grid origin)
    elements.push(
      <div
        key="margins"
        style={{
          position: "absolute",
          left: `${-originX + marginLeftPt}pt`,
          top: `${-originY + marginTopPt}pt`,
          width: `${paperW - marginLeftPt - marginRightPt}pt`,
          height: `${paperH - marginTopPt - marginBottomPt}pt`,
          border: "1px dashed rgba(0, 180, 0, 0.6)",
          zIndex: 4,
        }}
        title="Printable area (within margins)"
      />
    );
  }

  return <div style={overlayStyle}>{elements}</div>;
}
