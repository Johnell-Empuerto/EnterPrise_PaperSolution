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
 * Date field — renders <input type="date">.
 * Falls back gracefully to text input on browsers that don't support date inputs.
 */
export function DateField({ overlay, value, onChange, production }: FieldComponentProps) {
  return (
    <input
      type="date"
      className="runtime-field"
      value={(value as string) ?? ""}
      onChange={(e) => onChange(e.target.value || null)}
      style={inputStyle(overlay, production)}
    />
  );
}

function inputStyle(overlay: OverlayModel, production?: boolean): React.CSSProperties {
  const borderColor = production
    ? "rgba(234, 179, 8, 0.6)"       // yellow-500
    : "rgba(245, 158, 11, 0.5)";     // amber-500

  return {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    padding: "1px 3px",
    border: production ? `2px solid ${borderColor}` : `1px solid ${borderColor}`,
    borderRadius: "2px",
    background: production ? "rgba(254, 249, 195, 0.25)" : "rgba(255, 255, 255, 0.85)",
    fontFamily: "Calibri, sans-serif",
    fontSize: `11pt`,
    color: "#1a1a1a",
    outline: "none",
    cursor: "text",
    transition: "border-color 0.15s, box-shadow 0.15s",
  };
}
