"use client";

import { useMemo } from "react";
import type { RuntimeForm } from "@/types/runtime";
import type { OverlayModel, OverlayType } from "@/types/overlay";
import { RuntimeCanvas } from "./RuntimeCanvas";
import { PageSurface } from "./PageSurface";
import { BackgroundLayer } from "./BackgroundLayer";
import type { RuntimeState } from "./useRuntimeState";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:5090";

export interface RuntimeFormViewerProps {
  /** The runtime form from the COM backend (includes PNG URL + pixel coords) */
  runtimeForm: RuntimeForm;
  /** The runtime state hook */
  runtime: RuntimeState;
  /** Optional: ID of a field to highlight with a selection border */
  selectedFieldId?: string | null;
  /** Current page (sheet index) to display — defaults to 0 */
  currentPage?: number;
}

/**
 * PaperLess Production Runtime — COM-powered form viewer.
 *
 * Architecture (mirrors the legacy WPF rendering pipeline):
 *
 *   PageSurface (WPF Canvas equivalent — fixed pixel size)
 *   ├── BackgroundLayer (WPF Image, Stretch=None — PNG at native pixels)
 *   ├── FieldLayer (WPF overlay controls — fields at absolute pixel positions)
 *   ├── AnnotationLayer (future)
 *   ├── SelectionLayer (future)
 *   └── DebugLayer (future)
 *
 * Key design principles:
 * - SINGLE coordinate space: 2550×3299px (at 300 DPI for Letter)
 * - No responsive layout ever affects page geometry
 * - Only the viewport scrolls — the PageSurface never scales
 * - All layers share identical dimensions
 */
export function RuntimeFormViewer({ runtimeForm, runtime, selectedFieldId, currentPage = 0 }: RuntimeFormViewerProps) {
  const sheet = runtimeForm.sheets[currentPage];
  if (!sheet) return null;

  // Convert only the CURRENT sheet's fields to OverlayModel[] (optimization: skip other sheets)
  const overlaysForSheet = useMemo(() => {
    const result: OverlayModel[] = [];
    for (const field of sheet.fields) {
      const type = fieldDataTypeToOverlayType(field.dataType);
      if (!type) continue;
      result.push({
        id: field.id,
        type,
        cell: field.cellReference,
        leftPt: field.leftPx,
        topPt: field.topPx,
        widthPt: field.widthPx,
        heightPt: field.heightPx,
        rotation: 0,
        metadata: {
          sheet: sheet.name,
          mergeRange: field.mergeRange,
          dataType: field.dataType,
        },
      });
    }
    return result;
  }, [sheet]);

  // Resolve background image URL
  const bgUrl = sheet.backgroundImage
    ? `${API_BASE_URL}${sheet.backgroundImage.startsWith("/") ? "" : "/"}${sheet.backgroundImage}`
    : `${API_BASE_URL}/preview/page_${runtimeForm.title}.png`;

  const sheetCollection = {
    templateId: parseInt(runtimeForm.title, 10) || 0,
    overlays: overlaysForSheet,
    byId: Object.fromEntries(overlaysForSheet.map((o) => [o.id, o])),
    byCell: Object.fromEntries(overlaysForSheet.map((o) => [o.cell, o])),
    generatedAt: new Date().toISOString(),
  };

  const hasSheetFields = overlaysForSheet.length > 0;

  return (
    <>
      <PageSurface
        key={sheet.name + currentPage}
        widthPx={sheet.pageWidthPx}
        heightPx={sheet.pageHeightPx}
      >
        {/* Layer 1: Background PNG */}
        <BackgroundLayer
          src={bgUrl}
          alt={`${runtimeForm.workbookName} — ${sheet.name}`}
          widthPx={sheet.pageWidthPx}
          heightPx={sheet.pageHeightPx}
        />

        {/* Layer 2: Interactive runtime fields */}
        {hasSheetFields && (
          <RuntimeCanvas
            overlayCollection={sheetCollection}
            runtimeValues={runtime.values}
            onValueChange={runtime.setValue}
            widthPt={sheet.pageWidthPx}
            heightPt={sheet.pageHeightPx}
            production
            usePixelUnits
          />
        )}

        {/* Layer 3: Selection highlight */}
        {selectedFieldId && (() => {
          const selOverlay = sheetCollection.byId[selectedFieldId];
          if (!selOverlay) return null;
          return (
            <div
              style={{
                position: "absolute",
                left: selOverlay.leftPt,
                top: selOverlay.topPt,
                width: selOverlay.widthPt,
                height: selOverlay.heightPt,
                border: "2px solid #2196F3",
                backgroundColor: "rgba(33, 150, 243, 0.08)",
                pointerEvents: "none",
                zIndex: 30,
                boxSizing: "border-box",
              }}
            />
          );
        })()}
      </PageSurface>

      {/* Empty state */}
      {!hasSheetFields && (
        <div className="bg-amber-50 text-amber-700 text-sm px-4 py-2 rounded-lg border border-amber-200" data-runtime-empty>
          No editable fields detected in this template
        </div>
      )}
    </>
  );
}

/**
 * Maps backend RuntimeField dataType to OverlayType for RuntimeCanvas compatibility.
 * Returns null for non-interactive types.
 */
function fieldDataTypeToOverlayType(dataType: string): OverlayType | null {
  const map: Record<string, OverlayType> = {
    text: "textbox",
    number: "number",
    date: "date",
    checkbox: "checkbox",
    signature: "signature",
    dropdown: "textbox",       // Fall back to textbox for dropdown
    calculated: "textbox",     // Fall back to textbox for calculated
  };
  return map[dataType] ?? null;
}
