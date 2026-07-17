"use client";

import { useCallback, memo, useRef } from "react";
import type { FieldComponentProps } from "@/runtime/components/fields";
import {
  DEFAULTS,
  toCharacterRestriction,
  toHorizontalAlign,
  toVerticalAlign,
  toFontWeight,
  rgbStringToHex,
} from "@/runtime/config/keyboardTextConfig";
import { convertLegacyConfigToKtParams } from "@/runtime/config/migration";

const RESTRICTION_PATTERNS: Record<string, RegExp> = {
  letters: /^[A-Za-z]*$/,
  numbers: /^[0-9]*$/,
  alphanumeric: /^[A-Za-z0-9]*$/,
};

function KeyboardTextFieldInner({
  overlay,
  value: storeValue,
  onChange,
  onBlur,
  production,
  config,
  disabled,
  readOnly: propReadOnly,
  required: propRequired,
  placeholder,
}: FieldComponentProps) {
  const keyboardTextParams = (config as any)?.keyboardText;
  const kt = keyboardTextParams?.required !== undefined
    ? keyboardTextParams
    : config
    ? convertLegacyConfigToKtParams(config as any)
    : DEFAULTS;

  if (kt.hidden) return null;

  const isMultiline = kt.lines > 1;
  const readOnly = propReadOnly ?? kt.readOnly;
  const enabled = !(disabled ?? false);
  const maxLength = kt.maxLength > 0 ? kt.maxLength : undefined;
  const restriction = toCharacterRestriction(kt.inputRestriction);
  const required = propRequired ?? kt.required;
  const displayValue = (storeValue as string | undefined) ?? kt.defaultValue ?? "";

  const inputRef = useRef<HTMLTextAreaElement | HTMLInputElement>(null);

  const handleWrapperClick = useCallback(() => {
    inputRef.current?.focus();
  }, []);

  const fontFamily = kt.font || DEFAULTS.font;
  const fontSize = kt.fontSize || DEFAULTS.fontSize;
  const fontWeight = toFontWeight(kt.weight);
  const textColor = rgbStringToHex(kt.color);
  const hAlign = toHorizontalAlign(kt.align);
  const vAlign = toVerticalAlign(kt.verticalAlignment);

  const VERTICAL_ALIGN_MAP: Record<string, React.CSSProperties["alignItems"]> = {
    top: "flex-start",
    middle: "center",
    bottom: "flex-end",
  };

  const fieldStyle: React.CSSProperties = {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    display: "flex",
    alignItems: VERTICAL_ALIGN_MAP[vAlign] ?? "center",
    overflow: "hidden",
    border: production
      ? "2px solid rgba(234, 179, 8, 0.6)"
      : "1px solid rgba(59, 130, 246, 0.5)",
    borderRadius: "2px",
    background: config?.appearance?.backgroundColor ?? (production ? "rgba(254, 249, 195, 0.25)" : "rgba(255, 255, 255, 0.85)"),
    transition: "border-color 0.15s, box-shadow 0.15s",
    opacity: enabled ? 1 : 0.5,
    cursor: enabled ? "text" : "not-allowed",
    ...(required && !readOnly
      ? { borderLeft: "3px solid #dc2626" }
      : {}),
  };

  const innerStyle: React.CSSProperties = {
    width: "100%",
    boxSizing: "border-box",
    border: "none",
    background: "transparent",
    outline: "none",
    fontFamily,
    fontSize: `${fontSize}pt`,
    fontWeight,
    color: textColor,
    textAlign: hAlign as React.CSSProperties["textAlign"],
    padding: "1px 3px",
    resize: "none",
    overflow: "hidden",
  };

  const handleChange = useCallback(
    (raw: string) => {
      if (readOnly || !enabled) return;
      let val = raw;

      if (restriction && RESTRICTION_PATTERNS[restriction]) {
        const pattern = RESTRICTION_PATTERNS[restriction];
        val = val.split("").filter((c) => pattern.test(c)).join("");
      }

      if (maxLength !== undefined && val.length > maxLength) {
        val = val.slice(0, maxLength);
      }

      onChange(val || null);
    },
    [readOnly, enabled, restriction, maxLength, onChange],
  );

  if (readOnly) {
    return (
      <div
        style={{
          ...fieldStyle,
          background: config?.appearance?.backgroundColor ?? "#f1f5f9",
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
          {displayValue}
        </span>
      </div>
    );
  }

  if (isMultiline) {
    return (
      <div style={fieldStyle} onClick={handleWrapperClick}>
        <textarea
          ref={inputRef as React.RefObject<HTMLTextAreaElement>}
          value={displayValue}
          onChange={(e) => handleChange(e.target.value)}
          onBlur={onBlur}
          placeholder={placeholder ?? ""}
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
        type="text"
        value={displayValue}
        onChange={(e) => handleChange(e.target.value)}
        onBlur={onBlur}
        placeholder={placeholder ?? ""}
        maxLength={maxLength}
        readOnly={readOnly}
        disabled={!enabled}
        style={innerStyle}
        className="runtime-field runtime-input"
      />
    </div>
  );
}

export const KeyboardTextField = memo(KeyboardTextFieldInner);
