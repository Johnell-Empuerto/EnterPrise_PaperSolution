"use client";

import { useMemo } from "react";
import type { OverlayModel, OverlayCollection } from "@/types/overlay";
import { RuntimeField } from "./RuntimeField";

export interface RuntimeCanvasProps {
  overlayCollection: OverlayCollection;
  runtimeValues: Record<string, string | boolean | null>;
  onValueChange: (overlayId: string, value: string | boolean | null) => void;
  widthPt: number;
  heightPt: number;
  production?: boolean;
  usePixelUnits?: boolean;
}

const INTERACTIVE_TYPES = new Set(["KeyboardText", "checkbox", "date", "number", "signature"]);

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
    [overlayCollection],
  );

  if (interactiveOverlays.length === 0) return null;

  const unit = usePixelUnits ? "px" : "pt";
  const canvasStyle: React.CSSProperties = {
    position: "absolute",
    left: 0,
    top: 0,
    width: `${widthPt}${unit}`,
    height: `${heightPt}${unit}`,
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
