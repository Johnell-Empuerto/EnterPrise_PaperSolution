"use client";

import { useMemo } from "react";
import type { OverlayModel, OverlayCollection } from "@/types/overlay";
import { RuntimeField } from "./RuntimeField";

export interface RuntimeCanvasProps {
  overlayCollection: OverlayCollection;
  runtimeValues: Record<string, string | boolean | null>;
  onValueChange: (overlayId: string, value: string | boolean | null) => void;
  /** Width of the grid in pt (for sizing the canvas) */
  widthPt: number;
  /** Height of the grid in pt (for sizing the canvas) */
  heightPt: number;
  /** Production mode: use yellow styling for all fields */
  production?: boolean;
  /** Use px units instead of pt (for COM-based coordinate system where pixels are source of truth) */
  usePixelUnits?: boolean;
}

/** Overlay types that produce interactive form fields */
const INTERACTIVE_TYPES = new Set(["textbox", "checkbox", "date", "number", "signature"]);

export function RuntimeCanvas({
  overlayCollection,
  runtimeValues,
  onValueChange,
  widthPt,
  heightPt,
  production,
  usePixelUnits,
}: RuntimeCanvasProps) {
  const interactiveOverlays = useMemo(
    () => overlayCollection.overlays.filter((o) => INTERACTIVE_TYPES.has(o.type)),
    [overlayCollection]
  );

  if (interactiveOverlays.length === 0) return null;

  // Position the canvas at the grid origin (0,0), same size as the grid
  const unit = usePixelUnits ? "px" : "pt";
  const canvasStyle: React.CSSProperties = {
    position: "absolute",
    left: 0,
    top: 0,
    width: `${widthPt}${unit}`,
    height: `${heightPt}${unit}`,
    pointerEvents: "none", // Allow clicks to pass through to the grid
    zIndex: 25,
  };

  return (
    <div style={canvasStyle}>
      {interactiveOverlays.map((overlay) => (
        <RuntimeField
          key={overlay.id}
          overlay={overlay}
          value={runtimeValues[overlay.id] ?? null}
          onChange={(val) => onValueChange(overlay.id, val)}
          production={production}
          usePixelUnits={usePixelUnits}
        />
      ))}
    </div>
  );
}
