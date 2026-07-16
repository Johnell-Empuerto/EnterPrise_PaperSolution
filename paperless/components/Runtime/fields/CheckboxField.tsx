"use client";

import type { OverlayModel, OverlayConfig } from "@/types/overlay";

export interface FieldComponentProps {
  overlay: OverlayModel;
  value: string | boolean | null;
  onChange: (value: string | boolean | null) => void;
  onBlur?: () => void;
  production?: boolean;
  config?: OverlayConfig;
  disabled?: boolean;
  readOnly?: boolean;
}

export function CheckboxField({
  overlay,
  value,
  onChange,
  production,
  config,
  disabled,
  readOnly,
}: FieldComponentProps) {
  const behaviorCfg = config?.behavior ?? {};
  const enabled = !(disabled ?? !(behaviorCfg.enabled ?? true));
  const isReadOnly = readOnly ?? behaviorCfg.readOnly ?? false;

  const boxSize = Math.min(overlay.widthPt, overlay.heightPt, 20);

  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: isReadOnly || !enabled ? "default" : "pointer",
        opacity: enabled ? 1 : 0.5,
      }}
    >
      <input
        type="checkbox"
        className="runtime-field"
        checked={value === true}
        onChange={(e) => {
          if (isReadOnly || !enabled) return;
          onChange(e.target.checked);
        }}
        disabled={!enabled}
        style={{
          width: `${boxSize}pt`,
          height: `${boxSize}pt`,
          cursor: isReadOnly || !enabled ? "default" : "pointer",
          accentColor: production ? "#eab308" : "#3b82f6",
          outline: production ? "1px solid rgba(234, 179, 8, 0.4)" : undefined,
          outlineOffset: "1px",
        }}
      />
    </div>
  );
}
