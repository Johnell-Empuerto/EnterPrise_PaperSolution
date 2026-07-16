"use client";

import { useCallback, memo, useRef } from "react";
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
  const layoutCfg = config?.layout ?? {};

  // Vertical alignment mapping — single source of truth from Layout config
  const VERTICAL_ALIGN_MAP: Record<string, React.CSSProperties["alignItems"]> = {
    top: "flex-start",
    middle: "center",
    bottom: "flex-end",
  };
  const verticalAlignment = VERTICAL_ALIGN_MAP[layoutCfg.verticalAlign ?? "middle"] ?? "center";

  const keyboardType = inputCfg.keyboardType ?? "text";
  const kb = KEYBOARD_MAP[keyboardType] ?? KEYBOARD_MAP.text;

  const isMultiline = propReadOnly ?? behaviorCfg.multiline ?? overlay.heightPt > 30;
  const readOnly = propReadOnly ?? behaviorCfg.readOnly ?? false;
  const enabled = !(disabled ?? !(behaviorCfg.enabled ?? true));
  const maxLength = inputCfg.maxLength ?? undefined;
  const charRestriction = inputCfg.characterRestriction ?? undefined;

  // Ref to the inner input/textarea so the wrapper can forward clicks
  const inputRef = useRef<HTMLTextAreaElement | HTMLInputElement>(null);

  // Click handler: forward wrapper clicks to the actual input/textarea
  const handleWrapperClick = useCallback(() => {
    inputRef.current?.focus();
  }, []);

  // ── Visual wrapper: border, background, flex alignment ──
  const fieldStyle: React.CSSProperties = {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    display: "flex",
    alignItems: verticalAlignment,
    overflow: "hidden",
    border: production
      ? "2px solid rgba(234, 179, 8, 0.6)"
      : "1px solid rgba(59, 130, 246, 0.5)",
    borderRadius: "2px",
    background: appearanceCfg.backgroundColor ?? (production ? "rgba(254, 249, 195, 0.25)" : "rgba(255, 255, 255, 0.85)"),
    transition: "border-color 0.15s, box-shadow 0.15s",
    opacity: enabled ? 1 : 0.5,
    cursor: enabled ? "text" : "not-allowed",
    ...(behaviorCfg.required && !readOnly
      ? { borderLeft: "3px solid #dc2626" }
      : {}),
  };

  // ── Inner element: font, text alignment, natural sizing ──
  const innerStyle: React.CSSProperties = {
    width: "100%",
    boxSizing: "border-box",
    border: "none",
    background: "transparent",
    outline: "none",
    fontFamily: appearanceCfg.fontFamily ?? "Calibri, sans-serif",
    fontSize: appearanceCfg.fontSize ? `${appearanceCfg.fontSize}pt` : "11pt",
    fontWeight: appearanceCfg.fontWeight ?? "normal",
    color: appearanceCfg.textColor ?? "#1a1a1a",
    textAlign: (layoutCfg.horizontalAlign as React.CSSProperties["textAlign"]) ?? "left",
    padding: "1px 3px",
    resize: "none",
    overflow: "hidden",
  };

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

  if (readOnly) {
    return (
      <div
        style={{
          ...fieldStyle,
          background: appearanceCfg.backgroundColor ?? "#f1f5f9",
          cursor: "default",
        }}
      >
        <span
          style={{
            fontSize: innerStyle.fontSize,
            fontFamily: innerStyle.fontFamily,
            fontWeight: innerStyle.fontWeight,
            color: innerStyle.color,
            padding: innerStyle.padding,
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
      <div style={fieldStyle} onClick={handleWrapperClick}>
        <textarea
          ref={inputRef as React.RefObject<HTMLTextAreaElement>}
          value={(value as string) ?? ""}
          onChange={(e) => handleChange(e.target.value)}
          onBlur={onBlur}
          placeholder=""
          maxLength={maxLength}
          readOnly={readOnly}
          disabled={!enabled}
          style={{ ...innerStyle, maxHeight: "100%", overflowY: "auto" }}
          className="runtime-field runtime-textarea"
        />
      </div>
    );
  }

  return (
    <div style={fieldStyle} onClick={handleWrapperClick}>
      <input
        ref={inputRef as React.RefObject<HTMLInputElement>}
        type={kb.type}
        inputMode={kb.inputMode}
        value={(value as string) ?? ""}
        onChange={(e) => handleChange(e.target.value)}
        onBlur={onBlur}
        placeholder=""
        maxLength={maxLength}
        readOnly={readOnly}
        disabled={!enabled}
        style={innerStyle}
        className="runtime-field runtime-input"
      />
    </div>
  );
}

export const TextField = memo(TextFieldInner);
