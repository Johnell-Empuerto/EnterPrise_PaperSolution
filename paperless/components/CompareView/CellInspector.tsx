"use client";

import type { TemplateModel, TemplateCellStyle } from "@/types/template";
import { cellRef, getCellValue, getCellStyle, cumulativeColWidth, cumulativeRowHeight } from "../ExcelRenderer/helpers";

export interface CellInspectorProps {
  template: TemplateModel;
  selectedCol: number | null;
  selectedRow: number | null;
}

export function CellInspector({ template, selectedCol, selectedRow }: CellInspectorProps) {
  if (selectedCol === null || selectedRow === null) {
    return (
      <div className="text-xs text-gray-400 italic p-3 text-center border border-dashed border-gray-200 rounded">
        Click any cell in the renderer to inspect its properties
      </div>
    );
  }

  const { columnWidths, rowHeights, cellStyles, cellValues, hiddenColumns, hiddenRows } = template;
  const style: TemplateCellStyle = getCellStyle(selectedCol, selectedRow, cellStyles);
  const value = getCellValue(selectedCol, selectedRow, cellValues);
  const ref = cellRef(selectedCol, selectedRow);

  const colW = columnWidths[selectedCol - 1] ?? 0;
  const rowH = rowHeights[selectedRow - 1] ?? 15;

  const absX = cumulativeColWidth(selectedCol, columnWidths, hiddenColumns);
  const absY = cumulativeRowHeight(selectedRow, rowHeights, hiddenRows);

  const fontSize = style.fontSize ?? 11;
  const fontName = style.fontName ?? "Calibri";
  const fontWeight = style.bold ? `700 (bold)` : `400 (normal)`;
  const fontStyle = style.italic ? "italic" : "normal";
  const fontDeco = style.underline ? "underline" : "none";
  const fontColor = style.fontColor ?? "#000000";
  const fillColor = style.fillColor ?? "transparent";
  const hAlign = style.horizontalAlignment ?? "left";
  const vAlign = style.verticalAlignment ?? "middle";
  const wrapText = style.wrapText ? "Yes" : "No";
  const indent = style.indent ?? 0;

  const borderTop = style.borderTop ?? "—";
  const borderRight = style.borderRight ?? "—";
  const borderBottom = style.borderBottom ?? "—";
  const borderLeft = style.borderLeft ?? "—";

  return (
    <div className="text-xs">
      {/* Header */}
      <div className="flex items-center gap-2 mb-2 pb-2 border-b border-gray-200">
        <span className="font-bold text-sm font-mono text-blue-700">{ref}</span>
        <span className="text-gray-400">Column {selectedCol} · Row {selectedRow}</span>
      </div>

      {/* Grid */}
      <div className="grid grid-cols-[auto_1fr] gap-x-3 gap-y-1.5">
        {/* Dimensions */}
        <SectionLabel label="Dimensions" />
        <div />
        <PropRow name="Width" value={`${colW.toFixed(1)} pt`} />
        <PropRow name="Height" value={`${rowH.toFixed(1)} pt`} />

        {/* Position */}
        <SectionLabel label="Position" />
        <div />
        <PropRow name="Left" value={`${absX.toFixed(1)} pt`} />
        <PropRow name="Top" value={`${absY.toFixed(1)} pt`} />
        <PropRow name="Right" value={`${(absX + colW).toFixed(1)} pt`} />
        <PropRow name="Bottom" value={`${(absY + rowH).toFixed(1)} pt`} />

        {/* Font */}
        <SectionLabel label="Font" />
        <div />
        <PropRow name="Family" value={fontName} />
        <PropRow name="Size" value={`${fontSize} pt`} />
        <PropRow name="Weight" value={fontWeight} />
        <PropRow name="Style" value={fontStyle} />
        <PropRow name="Decoration" value={fontDeco} />
        <div className="flex items-center gap-2">
          <span className="text-gray-400">Color</span>
          <span
            style={{
              display: "inline-block",
              width: 12,
              height: 12,
              borderRadius: 2,
              background: fontColor,
              border: "1px solid #ccc",
              verticalAlign: "middle",
            }}
          />
          <span className="font-mono">{fontColor}</span>
        </div>

        {/* Fill */}
        <SectionLabel label="Fill" />
        <div />
        <div className="flex items-center gap-2">
          <span className="text-gray-400">Background</span>
          {fillColor !== "transparent" ? (
            <>
              <span
                style={{
                  display: "inline-block",
                  width: 12,
                  height: 12,
                  borderRadius: 2,
                  background: fillColor,
                  border: "1px solid #ccc",
                  verticalAlign: "middle",
                }}
              />
              <span className="font-mono">{fillColor}</span>
            </>
          ) : (
            <span className="font-mono text-gray-400">transparent</span>
          )}
        </div>

        {/* Alignment */}
        <SectionLabel label="Alignment" />
        <div />
        <PropRow name="Horizontal" value={hAlign} />
        <PropRow name="Vertical" value={vAlign} />
        <PropRow name="Wrap Text" value={wrapText} />
        <PropRow name="Indent" value={`${indent}`} />

        {/* Borders */}
        <SectionLabel label="Borders" />
        <div />
        <PropRow name="Top" value={borderTop} />
        <PropRow name="Right" value={borderRight} />
        <PropRow name="Bottom" value={borderBottom} />
        <PropRow name="Left" value={borderLeft} />
      </div>

      {/* Text Preview */}
      <div className="mt-2 pt-2 border-t border-gray-200">
        <span className="text-gray-400">Text</span>
        <div className="mt-0.5 font-mono text-xs bg-gray-50 p-1.5 rounded border border-gray-100 max-h-12 overflow-y-auto whitespace-pre-wrap break-words">
          {value || <span className="text-gray-300 italic">(empty)</span>}
        </div>
      </div>
    </div>
  );
}

function SectionLabel({ label }: { label: string }) {
  return (
    <div className="col-span-2 text-[10px] font-semibold text-gray-500 uppercase tracking-wider mt-1.5 first:mt-0">
      {label}
    </div>
  );
}

function PropRow({ name, value }: { name: string; value: string }) {
  return (
    <>
      <span className="text-gray-400">{name}</span>
      <span className="font-mono text-gray-700">{value}</span>
    </>
  );
}
