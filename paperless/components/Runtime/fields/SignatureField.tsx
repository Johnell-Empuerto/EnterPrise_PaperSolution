"use client";

import type { OverlayModel } from "@/types/overlay";

export interface FieldComponentProps {
  overlay: OverlayModel;
  value: string | boolean | null;
  onChange: (value: string | boolean | null) => void;
  production?: boolean;
  disabled?: boolean;
  readOnly?: boolean;
}

export function SignatureField({
  overlay,
  onChange,
  production,
  disabled,
  readOnly,
}: FieldComponentProps) {
  const isDisabled = disabled || readOnly;

  return (
    <div
      className="runtime-field"
      onClick={() => {
        if (!isDisabled) {
          onChange(null);
        }
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
        cursor: isDisabled ? "default" : "pointer",
        fontFamily: "Calibri, sans-serif",
        fontSize: `${Math.min(10, overlay.widthPt / 8)}pt`,
        color: production ? "#a16207" : "#7c3aed",
        userSelect: "none",
        opacity: isDisabled ? 0.5 : 1,
        transition: "background 0.15s",
      }}
      title="Click to Sign"
    >
      Sign
    </div>
  );
}
