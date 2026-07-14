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
 * Signature field — renders a placeholder "Click to Sign" button.
 * Actual signature drawing will be implemented in a later phase.
 */
export function SignatureField({ overlay, onChange, production }: FieldComponentProps) {
  return (
    <div
      className="runtime-field"
      onClick={() => {
        onChange(null); // Mark as interacted
      }}
      style={{
        width: "100%",
        height: "100%",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        border: production ? "2px dashed rgba(234, 179, 8, 0.6)" : "1.5px dashed rgba(139, 92, 246, 0.6)",
        borderRadius: "2px",
        background: production ? "rgba(254, 249, 195, 0.2)" : "rgba(139, 92, 246, 0.08)",
        cursor: "pointer",
        fontFamily: "Calibri, sans-serif",
        fontSize: `${Math.min(10, overlay.widthPt / 8)}pt`,
        color: production ? "#a16207" : "#7c3aed",
        userSelect: "none",
        transition: "background 0.15s",
      }}
      title="Click to Sign"
    >
      ✍ Sign
    </div>
  );
}
