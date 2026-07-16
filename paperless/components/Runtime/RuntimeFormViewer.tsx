"use client";

import { useMemo, useEffect } from "react";
import type { RuntimeForm } from "@/types/runtime";
import type { OverlayModel, OverlayType } from "@/types/overlay";
import { RuntimeCanvas } from "./RuntimeCanvas";
import { PageSurface } from "./PageSurface";
import { BackgroundLayer } from "./BackgroundLayer";
import type { RuntimeState } from "./useRuntimeState";
import { initializeStore, getDefaultStore } from "@/runtime/store";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:5090";

export interface RuntimeFormViewerProps {
  runtimeForm: RuntimeForm;
  runtime: RuntimeState;
  selectedFieldId?: string | null;
  currentPage?: number;
  showOverlay?: boolean;
  showBackground?: boolean;
}

export function RuntimeFormViewer({
  runtimeForm,
  runtime,
  selectedFieldId,
  currentPage = 0,
  showOverlay = true,
  showBackground = true,
}: RuntimeFormViewerProps) {
  // Initialize the Zustand store with form data on mount/data change
  useEffect(() => {
    if (runtimeForm) {
      const store = getDefaultStore();
      initializeStore(store, runtimeForm);
    }
  }, [runtimeForm]);

  const sheet = runtimeForm.sheets[currentPage];
  if (!sheet) return null;

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
          placeholder: field.placeholder,
        },
        config: field.config as any,
      });
    }
    return result;
  }, [sheet]);

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
        {showBackground && (
          <BackgroundLayer
            src={bgUrl}
            alt={`${runtimeForm.workbookName} — ${sheet.name}`}
            widthPx={sheet.pageWidthPx}
            heightPx={sheet.pageHeightPx}
          />
        )}

        {showOverlay && hasSheetFields && (
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

        {!showBackground && !showOverlay && (
          <div
            style={{
              position: "absolute",
              inset: 0,
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              backgroundColor: "#f8fafc",
              color: "#94a3b8",
              zIndex: 50,
            }}
          >
            <svg className="w-10 h-10 mb-3" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M3.98 8.223A10.477 10.477 0 001.934 12C3.226 16.338 7.244 19.5 12 19.5c.993 0 1.953-.138 2.863-.395M6.228 6.228A10.45 10.45 0 0112 4.5c4.756 0 8.773 3.162 10.065 7.498a10.523 10.523 0 01-4.293 5.774M6.228 6.228L3 3m3.228 3.228l3.65 3.65m7.894 7.894L21 21m-3.228-3.228l-3.65-3.65m0 0a3 3 0 10-4.243-4.243m4.242 4.242L9.88 9.88" />
            </svg>
            <p className="text-sm font-medium mb-2">Nothing is visible</p>
            <p className="text-xs text-center max-w-[200px]">
              Enable <strong>Fields</strong> or <strong>Background</strong> from the toolbar
            </p>
          </div>
        )}

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

      {!hasSheetFields && (
        <div className="bg-amber-50 text-amber-700 text-sm px-4 py-2 rounded-lg border border-amber-200" data-runtime-empty>
          No editable fields detected in this template
        </div>
      )}
    </>
  );
}

function fieldDataTypeToOverlayType(dataType: string): OverlayType | null {
  const map: Record<string, OverlayType> = {
    text: "textbox",
    number: "number",
    date: "date",
    checkbox: "checkbox",
    signature: "signature",
    dropdown: "textbox",
    calculated: "textbox",
  };
  return map[dataType] ?? null;
}
