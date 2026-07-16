"use client";

import { useCallback, memo } from "react";
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
  required?: boolean;
  placeholder?: string;
}

const KEYBOARD_MAP: Record<string, { type: string; inputMode?: React.HTMLAttributes<HTMLInputElement>["inputMode"] }> = {
  text: { type: "text" },
  number: { type: "number" },
  decimal: { type: "text", inputMode: "decimal" },
  email: { type: "email" },
  phone: { type: "tel" },
  password: { type: "password" },
  url: { type: "url" },
};

const RESTRICTION_PATTERNS: Record<string, RegExp> = {
  letters: /^[A-Za-z]*$/,
  numbers: /^[0-9]*$/,
  alphanumeric: /^[A-Za-z0-9]*$/,
  uppercase: /^[A-Z]*$/,
  lowercase: /^[a-z]*$/,
};

function TextFieldInner({
  overlay,
  value,
  onChange,
  onBlur,
  production,
  config,
  disabled,
  readOnly: propReadOnly,
}: FieldComponentProps) {
  const inputCfg = config?.input ?? {};
  const behaviorCfg = config?.behavior ?? {};
  const appearanceCfg = config?.appearance ?? {};

  const keyboardType = inputCfg.keyboardType ?? "text";
  const kb = KEYBOARD_MAP[keyboardType] ?? KEYBOARD_MAP.text;

  const isMultiline = propReadOnly ?? behaviorCfg.multiline ?? overlay.heightPt > 30;
  const readOnly = propReadOnly ?? behaviorCfg.readOnly ?? false;
  const enabled = !(disabled ?? !(behaviorCfg.enabled ?? true));
  const maxLength = inputCfg.maxLength ?? undefined;
  const charRestriction = inputCfg.characterRestriction ?? undefined;

  const handleChange = useCallback(
    (raw: string) => {
      if (readOnly || !enabled) return;

      let val = raw;

      if (charRestriction && RESTRICTION_PATTERNS[charRestriction]) {
        const pattern = RESTRICTION_PATTERNS[charRestriction];
        val = val.split("").filter((c) => pattern.test(c)).join("");
      }

      if (maxLength !== undefined && val.length > maxLength) {
        val = val.slice(0, maxLength);
      }

      if (charRestriction === "uppercase") val = val.toUpperCase();
      if (charRestriction === "lowercase") val = val.toLowerCase();

      onChange(val || null);
    },
    [readOnly, enabled, charRestriction, maxLength, onChange],
  );

  const style: React.CSSProperties = {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    padding: "1px 3px",
    border: appearanceCfg.border
      ? `1px ${appearanceCfg.border} ${production ? "rgba(234, 179, 8, 0.6)" : "rgba(59, 130, 246, 0.5)"}`
      : production
        ? "2px solid rgba(234, 179, 8, 0.6)"
        : "1px solid rgba(59, 130, 246, 0.5)",
    borderRadius: appearanceCfg.borderRadius ?? "2px",
    background: appearanceCfg.backgroundColor ?? (production ? "rgba(254, 249, 195, 0.25)" : "rgba(255, 255, 255, 0.85)"),
    fontFamily: appearanceCfg.fontFamily ?? "Calibri, sans-serif",
    fontSize: appearanceCfg.fontSize ? `${appearanceCfg.fontSize}pt` : "11pt",
    fontWeight: appearanceCfg.fontWeight ?? "normal",
    color: appearanceCfg.textColor ?? "#1a1a1a",
    textAlign: (appearanceCfg.textAlign as any) ?? "left",
    outline: "none",
    transition: "border-color 0.15s, box-shadow 0.15s",
    resize: "none",
    overflow: "hidden",
    opacity: enabled ? 1 : 0.5,
    cursor: enabled ? undefined : "not-allowed",
    ...(behaviorCfg.required && !readOnly
      ? { borderLeft: "3px solid #dc2626" }
      : {}),
  };

  if (readOnly) {
    return (
      <div
        style={{
          ...style,
          display: "flex",
          alignItems: "center",
          background: appearanceCfg.backgroundColor ?? "#f1f5f9",
          cursor: "default",
        }}
      >
        <span
          style={{
            fontSize: style.fontSize,
            fontFamily: style.fontFamily,
            fontWeight: style.fontWeight,
            color: style.color,
            opacity: 0.6,
          }}
        >
          {(value as string) || ""}
        </span>
      </div>
    );
  }

  if (isMultiline) {
    return (
      <textarea
        value={(value as string) ?? ""}
        onChange={(e) => handleChange(e.target.value)}
        onBlur={onBlur}
        placeholder=""
        maxLength={maxLength}
        readOnly={readOnly}
        disabled={!enabled}
        style={style}
        className="runtime-field runtime-textarea"
      />
    );
  }

  return (
    <input
      type={kb.type}
      inputMode={kb.inputMode}
      value={(value as string) ?? ""}
      onChange={(e) => handleChange(e.target.value)}
      onBlur={onBlur}
      placeholder=""
      maxLength={maxLength}
      readOnly={readOnly}
      disabled={!enabled}
      style={style}
      className="runtime-field runtime-input"
    />
  );
}

export const TextField = memo(TextFieldInner);
