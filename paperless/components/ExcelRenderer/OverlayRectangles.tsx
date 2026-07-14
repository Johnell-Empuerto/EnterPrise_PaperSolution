"use client";

import { useMemo } from "react";
import type { TemplateModel } from "@/types/template";
import type { OverlayModel, OverlayType } from "@/types/overlay";
import { generateOverlays } from "@/services/overlayEngine";

export interface OverlayRectanglesProps {
  template: TemplateModel;
  showOverlays: boolean;
  onOverlayClick?: (overlay: OverlayModel) => void;
}

/** Color map for overlay types */
const OVERLAY_COLORS: Record<OverlayType, string> = {
  textbox: "rgba(59, 130, 246, 0.25)",   // blue
  signature: "rgba(139, 92, 246, 0.25)",  // purple
  checkbox: "rgba(16, 185, 129, 0.25)",   // green
  date: "rgba(245, 158, 11, 0.25)",       // amber
  number: "rgba(239, 68, 68, 0.25)",      // red
  qr: "rgba(236, 72, 153, 0.25)",         // pink
  barcode: "rgba(249, 115, 22, 0.25)",    // orange
  image: "rgba(168, 85, 247, 0.25)",      // violet
  ocr: "rgba(20, 184, 166, 0.25)",        // teal
  unknown: "rgba(156, 163, 175, 0.25)",   // gray
};

const OVERLAY_BORDER_COLORS: Record<OverlayType, string> = {
  textbox: "rgba(59, 130, 246, 0.7)",
  signature: "rgba(139, 92, 246, 0.7)",
  checkbox: "rgba(16, 185, 129, 0.7)",
  date: "rgba(245, 158, 11, 0.7)",
  number: "rgba(239, 68, 68, 0.7)",
  qr: "rgba(236, 72, 153, 0.7)",
  barcode: "rgba(249, 115, 22, 0.7)",
  image: "rgba(168, 85, 247, 0.7)",
  ocr: "rgba(20, 184, 166, 0.7)",
  unknown: "rgba(156, 163, 175, 0.7)",
};

const OVERLAY_LABELS: Record<OverlayType, string> = {
  textbox: "TEXT",
  signature: "SIG",
  checkbox: "CHK",
  date: "DATE",
  number: "NUM",
  qr: "QR",
  barcode: "BAR",
  image: "IMG",
  ocr: "OCR",
  unknown: "???",
};

export function OverlayRectangles({ template, showOverlays, onOverlayClick }: OverlayRectanglesProps) {
  const collection = useMemo(() => generateOverlays(template), [template]);

  if (!showOverlays || collection.overlays.length === 0) return null;

  const containerStyle: React.CSSProperties = {
    position: "absolute",
    inset: 0,
    pointerEvents: "none",
    userSelect: "none",
    zIndex: 15,
  };

  return (
    <div style={containerStyle}>
      {collection.overlays.map((overlay, i) => {
        const bg = OVERLAY_COLORS[overlay.type] ?? OVERLAY_COLORS.unknown;
        const border = OVERLAY_BORDER_COLORS[overlay.type] ?? OVERLAY_BORDER_COLORS.unknown;
        const label = OVERLAY_LABELS[overlay.type] ?? "???";

        return (
          <div
            key={overlay.id}
            style={{
              position: "absolute",
              left: `${overlay.leftPt}pt`,
              top: `${overlay.topPt}pt`,
              width: `${overlay.widthPt}pt`,
              height: `${overlay.heightPt}pt`,
              background: bg,
              border: `1.5px solid ${border}`,
              borderRadius: "1px",
              cursor: onOverlayClick ? "pointer" : undefined,
              pointerEvents: onOverlayClick ? "auto" : "none",
              transition: "background 0.15s, border-color 0.15s",
            }}
            onClick={(e) => {
              if (onOverlayClick) {
                e.stopPropagation();
                onOverlayClick(overlay);
              }
            }}
            title={`${label}: ${overlay.cell} (${overlay.widthPt.toFixed(1)}×${overlay.heightPt.toFixed(1)}pt)`}
          >
            {/* Type label badge */}
            <span
              style={{
                position: "absolute",
                top: 0,
                left: 0,
                fontSize: "6pt",
                fontFamily: "monospace",
                fontWeight: 700,
                color: border,
                background: "rgba(255,255,255,0.8)",
                padding: "0 2px",
                lineHeight: "10pt",
                borderRadius: "0 0 2px 0",
              }}
            >
              {label}
            </span>
            {/* Cell reference badge */}
            <span
              style={{
                position: "absolute",
                bottom: 0,
                right: 0,
                fontSize: "6pt",
                fontFamily: "monospace",
                color: border,
                background: "rgba(255,255,255,0.8)",
                padding: "0 2px",
                lineHeight: "10pt",
                borderRadius: "2px 0 0 0",
              }}
            >
              {overlay.cell}
            </span>
          </div>
        );
      })}
    </div>
  );
}
