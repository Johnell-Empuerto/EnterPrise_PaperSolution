"use client";

import type { RuntimeField } from "@/types/runtime";
import { OverlayField } from "./OverlayField";

interface OverlayRendererProps {
  /** All fields to render as overlays. */
  fields: RuntimeField[];
  /** Called when a field is focused. */
  onFieldFocus?: (field: RuntimeField) => void;
  /** Current zoom level (passed to fields for coordinate scaling). */
  scale: number;
  /** Page width in pixels for ratio-based rendering fallback. */
  pageWidthPx: number;
  /** Page height in pixels for ratio-based rendering fallback. */
  pageHeightPx: number;
  /** Visual debug mode: renders green border around each overlay field. */
  debug?: boolean;
}

/**
 * Renders all editable overlay fields positioned absolutely on top of the background.
 * Every field uses coordinates from the backend RuntimeField (pixel or ratio-based).
 * Positions are never recalculated in the frontend.
 */
export function OverlayRenderer({ fields, onFieldFocus, scale, pageWidthPx, pageHeightPx, debug }: OverlayRendererProps) {
  // ── DEBUG: Log overlay render info ──────────────────────────────────
  console.log("[DEBUG] OverlayRenderer:", {
    fieldCount: fields?.length ?? 0,
    pageWidthPx,
    pageHeightPx,
    scale,
    debug,
    firstField: fields?.[0]
      ? { cell: fields[0].cellReference, leftPx: fields[0].leftPx, topPx: fields[0].topPx, leftRatio: fields[0].leftRatio }
      : null,
  });

  if (!fields || fields.length === 0) {
    return (
      <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
        <p className="text-sm text-gray-400 pointer-events-auto bg-white/80 px-4 py-2 rounded-lg shadow-sm">
          No editable fields detected
        </p>
      </div>
    );
  }

  return (
    <>
      {fields.map((field) => (
        <OverlayField
          key={field.id}
          field={field}
          pageWidthPx={pageWidthPx}
          pageHeightPx={pageHeightPx}
          onFocus={() => onFieldFocus?.(field)}
          debug={debug}
        />
      ))}
    </>
  );
}
