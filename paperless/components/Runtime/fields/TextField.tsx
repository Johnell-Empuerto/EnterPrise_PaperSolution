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
 * Text input field — renders <input type="text"> or <textarea> for large fields.
 */
export function TextField({ overlay, value, onChange, production }: FieldComponentProps) {
  const isLarge = overlay.heightPt > 30;
  const style = inputStyle(overlay, production);

  if (isLarge) {
    return (
      <textarea
        value={(value as string) ?? ""}
        onChange={(e) => onChange(e.target.value || null)}
        placeholder=""
        style={style}
        className="runtime-field runtime-textarea"
      />
    );
  }

  return (
    <input
      type="text"
      value={(value as string) ?? ""}
      onChange={(e) => onChange(e.target.value || null)}
      placeholder=""
      style={style}
      className="runtime-field runtime-input"
    />
  );
}

function inputStyle(overlay: OverlayModel, production?: boolean): React.CSSProperties {
  const borderColor = production
    ? "rgba(234, 179, 8, 0.6)"      // yellow-500
    : "rgba(59, 130, 246, 0.5)";    // blue-500

  return {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    padding: "1px 3px",
    border: production ? `2px solid ${borderColor}` : `1px solid ${borderColor}`,
    borderRadius: "2px",
    background: production ? "rgba(254, 249, 195, 0.25)" : "rgba(255, 255, 255, 0.85)", // yellow-50 tint
    fontFamily: "Calibri, sans-serif",
    fontSize: `11pt`,
    color: "#1a1a1a",
    outline: "none",
    transition: "border-color 0.15s, box-shadow 0.15s",
    resize: "none",
    overflow: "hidden",
  };
}
