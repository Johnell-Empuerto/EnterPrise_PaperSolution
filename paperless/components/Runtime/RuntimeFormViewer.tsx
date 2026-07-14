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
export function RuntimeFormViewer({ runtimeForm, runtime }: RuntimeFormViewerProps) {
  // Convert backend RuntimeField[] to OverlayModel[] for RuntimeCanvas compatibility
  const overlays = useMemo(() => {
    const result: OverlayModel[] = [];
    for (const sheet of runtimeForm.sheets) {
      for (const field of sheet.fields) {
        // Map backend dataType to OverlayType
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
    }
    return result;
  }, [runtimeForm]);

  const overlaysBySheet = useMemo(() => {
    const map = new Map<string, OverlayModel[]>();
    for (const overlay of overlays) {
      const sheet = (overlay.metadata?.sheet as string) ?? "";
      if (!map.has(sheet)) map.set(sheet, []);
      map.get(sheet)!.push(overlay);
    }
    return map;
  }, [overlays]);

  const totalFields = overlays.length;

  return (
    <>
      {runtimeForm.sheets.map((sheet, idx) => {
        // Resolve background image URL
        const bgUrl = sheet.backgroundImage
          ? `${API_BASE_URL}${sheet.backgroundImage.startsWith("/") ? "" : "/"}${sheet.backgroundImage}`
          : `${API_BASE_URL}/preview/page_${runtimeForm.title}.png`;

        // Filter overlays to only this sheet (multi-sheet safety)
        const sheetOverlays = overlaysBySheet.get(sheet.name) ?? [];
        const sheetCollection = {
          templateId: parseInt(runtimeForm.title, 10) || 0,
          overlays: sheetOverlays,
          byId: Object.fromEntries(sheetOverlays.map((o) => [o.id, o])),
          byCell: Object.fromEntries(sheetOverlays.map((o) => [o.cell, o])),
          generatedAt: new Date().toISOString(),
        };
        const hasSheetFields = sheetOverlays.length > 0;

        return (
          <PageSurface
            key={sheet.name + idx}
            widthPx={sheet.pageWidthPx}
            heightPx={sheet.pageHeightPx}
          >
            {/* Layer 1: Background PNG — rendered by Excel via COM, displayed at native pixels */}
            <BackgroundLayer
              src={bgUrl}
              alt={`${runtimeForm.workbookName} — ${sheet.name}`}
              widthPx={sheet.pageWidthPx}
              heightPx={sheet.pageHeightPx}
            />

            {/* Layer 2: Interactive runtime fields — positioned over the PNG using pixel coords */}
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
          </PageSurface>
        );
      })}

      {/* Empty state */}
      {totalFields === 0 && runtimeForm.sheets.length > 0 && (
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
