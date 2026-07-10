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
}

/**
 * Renders all editable overlay fields positioned absolutely on top of the background.
 * Every field uses LeftPx/TopPx/WidthPx/HeightPx from the backend RuntimeField.
 * Positions are never recalculated in the frontend.
 */
export function OverlayRenderer({ fields, onFieldFocus, scale }: OverlayRendererProps) {
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
          onFocus={() => onFieldFocus?.(field)}
        />
      ))}
    </>
  );
}
