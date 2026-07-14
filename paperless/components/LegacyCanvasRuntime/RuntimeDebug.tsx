"use client";

import type { DebugFlags } from "./LegacyCanvasRuntime";

interface RuntimeDebugProps {
  flags: DebugFlags;
  onChange: (flags: DebugFlags) => void;
  fieldCount: number;
  pageWidth: number;
  pageHeight: number;
}

export default function RuntimeDebug({
  flags,
  onChange,
  fieldCount,
  pageWidth,
  pageHeight,
}: RuntimeDebugProps) {
  const toggle = (key: keyof DebugFlags) => {
    onChange({ ...flags, [key]: !flags[key] });
  };

  const switches: { key: keyof DebugFlags; label: string }[] = [
    { key: "showPageBounds", label: "Page bounds" },
    { key: "showFieldRectangles", label: "Field rectangles" },
    { key: "showFieldIds", label: "Field IDs" },
    { key: "showCoordinates", label: "Coordinates" },
    { key: "showOriginMarker", label: "Origin marker" },
    { key: "showPixelGrid", label: "Pixel grid" },
    { key: "showHitTestRegions", label: "Hit-test regions" },
  ];

  return (
    <div
      style={{
        position: "fixed",
        bottom: 12,
        right: 12,
        background: "rgba(30,30,30,0.92)",
        color: "#ccc",
        borderRadius: 8,
        padding: "10px 14px",
        font: "11px monospace",
        zIndex: 100,
        minWidth: 180,
        boxShadow: "0 2px 12px rgba(0,0,0,0.4)",
      }}
    >
      <div style={{ fontWeight: "bold", marginBottom: 6, color: "#eee", fontSize: 12 }}>
        Debug
      </div>
      {switches.map((s) => (
        <label
          key={s.key}
          style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", marginBottom: 3 }}
        >
          <input
            type="checkbox"
            checked={flags[s.key]}
            onChange={() => toggle(s.key)}
          />
          {s.label}
        </label>
      ))}
      <div style={{ marginTop: 6, borderTop: "1px solid #555", paddingTop: 6, fontSize: 10, color: "#999" }}>
        {fieldCount} fields &middot; {pageWidth}×{pageHeight}px
      </div>
    </div>
  );
}
