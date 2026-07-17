/**
 * Phase 28 — Overlay Coordinate Engine
 *
 * Pure mathematical coordinate calculation engine.
 * Converts Excel comments + template data into pixel-accurate overlay definitions.
 *
 * No DOM queries. No screenshots. Pure calculation from columnWidths, rowHeights,
 * mergedCells, page margins, and print origin.
 *
 * Performance target: <10ms normal templates, <50ms very large templates.
 */

import type { TemplateModel, TemplateComment, TemplateMergedCell, TemplateImage } from "@/types/template";
import type { OverlayModel, OverlayType, OverlayCollection } from "@/types/overlay";
import { parseCellRef, cellRef, cumulativeColWidth, cumulativeRowHeight } from "@/components/ExcelRenderer/helpers";

/**
 * Parses the comment text to determine the overlay type.
 * Comments can contain type hints like "textbox", "signature", "checkbox", etc.
 * Falls back to "textbox" for most fields, "unknown" if unparseable.
 */
function inferOverlayType(commentText: string): OverlayType {
  const lower = commentText.toLowerCase().trim();

  const typeMap: Record<string, OverlayType> = {
    textbox: "KeyboardText",
    text: "KeyboardText",
    "text field": "KeyboardText",
    input: "KeyboardText",
    keyboard: "KeyboardText",
    signature: "signature",
    "sign here": "signature",
    checkbox: "checkbox",
    "check box": "checkbox",
    tick: "checkbox",
    date: "date",
    "date field": "date",
    number: "number",
    "number field": "number",
    amount: "number",
    price: "number",
    qr: "qr",
    "qr code": "qr",
    barcode: "barcode",
    "bar code": "barcode",
    image: "image",
    picture: "image",
    photo: "image",
    ocr: "ocr",
    "ocr region": "ocr",
    "scan region": "ocr",
  };

  // Check for exact match first
  if (typeMap[lower]) return typeMap[lower];

  // Check if any keyword is contained in the text
  for (const [keyword, type] of Object.entries(typeMap)) {
    if (lower.includes(keyword)) return type;
  }

  return "KeyboardText"; // Default for form fields
}

/**
 * Gets the full span rectangle for a merged cell.
 * If the cell is part of a merge, returns the full merge bounds.
 * If not merged, returns the single cell bounds.
 */
function getCellRect(
  col: number,
  row: number,
  columnWidths: number[],
  rowHeights: number[],
  mergedCells: TemplateMergedCell[],
  hiddenColumns?: boolean[],
  hiddenRows?: boolean[]
): { leftPt: number; topPt: number; widthPt: number; heightPt: number } {
  // Check if this cell is inside any merge
  for (const mc of mergedCells) {
    if (col >= mc.startCol && col <= mc.endCol && row >= mc.startRow && row <= mc.endRow) {
      // Return the full merge rectangle
      const leftPt = cumulativeColWidth(mc.startCol, columnWidths, hiddenColumns);
      const topPt = cumulativeRowHeight(mc.startRow, rowHeights, hiddenRows);
      const rightPt = cumulativeColWidth(mc.endCol + 1, columnWidths, hiddenColumns);
      const bottomPt = cumulativeRowHeight(mc.endRow + 1, rowHeights, hiddenRows);
      return {
        leftPt,
        topPt,
        widthPt: rightPt - leftPt,
        heightPt: bottomPt - topPt,
      };
    }
  }

  // Single cell
  const leftPt = cumulativeColWidth(col, columnWidths, hiddenColumns);
  const topPt = cumulativeRowHeight(row, rowHeights, hiddenRows);
  return {
    leftPt,
    topPt,
    widthPt: columnWidths[col - 1] ?? 0,
    heightPt: rowHeights[row - 1] ?? 15,
  };
}

/**
 * Generates all overlays from a TemplateModel.
 *
 * Sources:
 * - Comments → form field overlays
 * - Merged cells → merged region overlays
 * - Images → image overlays
 *
 * @param template The template model
 * @returns An OverlayCollection with overlays indexed by id and cell
 */
export function generateOverlays(template: TemplateModel): OverlayCollection {
  const { comments, mergedCells, columnWidths, rowHeights, hiddenColumns, hiddenRows, images } = template;
  const overlays: OverlayModel[] = [];
  const byId: Record<string, OverlayModel> = {};
  const byCell: Record<string, OverlayModel> = {};

  // ── 1. Generate overlays from comments ──
  let commentIndex = 0;
  for (const comment of comments) {
    const parsed = parseCellRef(comment.address);
    if (!parsed) continue;

    const { col, row } = parsed;
    const rect = getCellRect(col, row, columnWidths, rowHeights, mergedCells, hiddenColumns, hiddenRows);
    const type = inferOverlayType(comment.text);

    const id = `field_${comment.address}`;
    const overlay: OverlayModel = {
      id,
      type,
      cell: comment.address,
      leftPt: rect.leftPt,
      topPt: rect.topPt,
      widthPt: rect.widthPt,
      heightPt: rect.heightPt,
      rotation: 0,
      metadata: {
        commentText: comment.text,
        commentIndex,
        source: "comment",
        col,
        row,
      },
    };

    overlays.push(overlay);
    byId[id] = overlay;
    byCell[comment.address] = overlay;
    commentIndex++;
  }

  // ── 2. Generate overlays from merged cells (if not already covered by comments) ──
  for (const mc of mergedCells) {
    const ref = cellRef(mc.startCol, mc.startRow);
    if (byCell[ref]) continue; // Already has a comment overlay

    const id = `merge_${ref}`;
    const overlay: OverlayModel = {
      id,
      type: "ocr",
      cell: `${ref}:${cellRef(mc.endCol, mc.endRow)}`,
      leftPt: cumulativeColWidth(mc.startCol, columnWidths, hiddenColumns),
      topPt: cumulativeRowHeight(mc.startRow, rowHeights, hiddenRows),
      widthPt: cumulativeColWidth(mc.endCol + 1, columnWidths, hiddenColumns) - cumulativeColWidth(mc.startCol, columnWidths, hiddenColumns),
      heightPt: cumulativeRowHeight(mc.endRow + 1, rowHeights, hiddenRows) - cumulativeRowHeight(mc.startRow, rowHeights, hiddenRows),
      rotation: 0,
      metadata: {
        source: "merge",
        startCol: mc.startCol,
        startRow: mc.startRow,
        endCol: mc.endCol,
        endRow: mc.endRow,
      },
    };

    overlays.push(overlay);
    byId[id] = overlay;
    byCell[ref] = overlay;
  }

  // ── 3. Generate overlays from images ──
  if (images) {
    for (let i = 0; i < images.length; i++) {
      const img = images[i];
      const ref = cellRef(img.anchorCol, img.anchorRow);

      const id = `image_${i}`;
      const overlay: OverlayModel = {
        id,
        type: "image",
        cell: ref,
        leftPt: cumulativeColWidth(img.anchorCol, columnWidths, hiddenColumns) + (img.offsetXPt ?? 0),
        topPt: cumulativeRowHeight(img.anchorRow, rowHeights, hiddenRows) + (img.offsetYPt ?? 0),
        widthPt: img.widthPt,
        heightPt: img.heightPt,
        rotation: 0,
        metadata: {
          source: "image",
          imageIndex: i,
          anchorCol: img.anchorCol,
          anchorRow: img.anchorRow,
          description: img.description,
        },
      };

      overlays.push(overlay);
      byId[id] = overlay;
      if (!byCell[ref]) byCell[ref] = overlay;
    }
  }

  return {
    templateId: 0,
    overlays,
    byId,
    byCell,
    generatedAt: new Date().toISOString(),
  };
}

/**
 * Look up an overlay by its cell reference.
 * Returns null if not found.
 */
export function findOverlayByCell(collection: OverlayCollection, cell: string): OverlayModel | null {
  return collection.byCell[cell.toUpperCase()] ?? null;
}

/**
 * Look up an overlay by its id.
 * Returns null if not found.
 */
export function findOverlayById(collection: OverlayCollection, id: string): OverlayModel | null {
  return collection.byId[id] ?? null;
}

/**
 * Export overlays as a JSON-serializable array.
 */
export function exportOverlays(collection: OverlayCollection): string {
  return JSON.stringify(collection.overlays, null, 2);
}

/**
 * Export overlays as a concise CSV-style format.
 */
export function exportOverlaysBrief(collection: OverlayCollection): string {
  const header = "id,type,cell,leftPt,topPt,widthPt,heightPt,rotation";
  const rows = collection.overlays.map(
    (o) => `${o.id},${o.type},${o.cell},${o.leftPt.toFixed(2)},${o.topPt.toFixed(2)},${o.widthPt.toFixed(2)},${o.heightPt.toFixed(2)},${o.rotation}`
  );
  return [header, ...rows].join("\n");
}
