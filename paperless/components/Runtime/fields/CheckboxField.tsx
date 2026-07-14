"use client";

import type { OverlayModel } from "@/types/overlay";

export interface FieldComponentProps {
  overlay: OverlayModel;
  value: string | boolean | null;
  onChange: (value: string | boolean | null) => void;
  /** Production mode: use yellow theme */
  production?: boolean;
}

/**
 * Checkbox field — renders <input type="checkbox"> centered inside the overlay.
 */
export function CheckboxField({ overlay, value, onChange, production }: FieldComponentProps) {
  const boxSize = Math.min(overlay.widthPt, overlay.heightPt, 20);

  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: "pointer",
      }}
      onClick={() => onChange(!(value === true))}
    >
      <input
        type="checkbox"
        className="runtime-field"
        checked={value === true}
        onChange={(e) => onChange(e.target.checked)}
        onClick={(e) => e.stopPropagation()}
        style={{
          width: `${boxSize}pt`,
          height: `${boxSize}pt`,
          cursor: "pointer",
          accentColor: production ? "#eab308" : "#3b82f6",
          outline: production ? "1px solid rgba(234, 179, 8, 0.4)" : undefined,
          outlineOffset: "1px",
        }}
      />
    </div>
  );
}
