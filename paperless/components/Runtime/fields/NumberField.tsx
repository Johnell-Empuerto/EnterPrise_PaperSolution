"use client";

import type { OverlayModel } from "@/types/overlay";

export interface FieldComponentProps {
  overlay: OverlayModel;
  value: string | boolean | null;
  onChange: (value: string | boolean | null) => void;
  onBlur?: () => void;
  production?: boolean;
  disabled?: boolean;
  readOnly?: boolean;
}

export function NumberField({
  overlay,
  value,
  onChange,
  onBlur,
  production,
  disabled,
  readOnly,
}: FieldComponentProps) {
  return (
    <input
      type="number"
      className="runtime-field"
      value={(value as string) ?? ""}
      onChange={(e) => onChange(e.target.value || null)}
      onBlur={onBlur}
      placeholder=""
      disabled={disabled}
      readOnly={readOnly}
      style={inputStyle(overlay, production)}
    />
  );
}

function inputStyle(overlay: OverlayModel, production?: boolean): React.CSSProperties {
  const borderColor = production
    ? "rgba(234, 179, 8, 0.6)"
    : "rgba(239, 68, 68, 0.5)";

  return {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    padding: "1px 3px",
    border: production ? `2px solid ${borderColor}` : `1px solid ${borderColor}`,
    borderRadius: "2px",
    background: production ? "rgba(254, 249, 195, 0.25)" : "rgba(255, 255, 255, 0.85)",
    fontFamily: "Calibri, sans-serif",
    fontSize: "11pt",
    color: "#1a1a1a",
    textAlign: "right" as const,
    outline: "none",
    transition: "border-color 0.15s, box-shadow 0.15s",
  };
}
