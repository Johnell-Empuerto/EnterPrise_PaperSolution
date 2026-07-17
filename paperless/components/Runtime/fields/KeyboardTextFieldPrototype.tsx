"use client";

import { useCallback, memo, useRef, useLayoutEffect } from "react";
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

function KeyboardTextFieldPrototypeInner({
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

  const maxLines = Math.max(1, kt.lines);
  const singleLineMode = maxLines === 1;
  const readOnly = propReadOnly ?? kt.readOnly;
  const enabled = !(disabled ?? false);
  const maxLength = kt.maxLength > 0 ? kt.maxLength : undefined;
  const restriction = toCharacterRestriction(kt.inputRestriction);
  const required = propRequired ?? kt.required;
  const displayValue = (storeValue as string | undefined) ?? kt.defaultValue ?? "";

  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const fontFamily = kt.font || DEFAULTS.font;
  const fontSize = kt.fontSize || DEFAULTS.fontSize;
  const fontWeight = toFontWeight(kt.weight);
  const textColor = rgbStringToHex(kt.color);
  const hAlign = toHorizontalAlign(kt.align);
  const vAlign = toVerticalAlign(kt.verticalAlignment);

  const LINE_HEIGHT_RATIO = 1.4;

  const handleChange = useCallback(
    (raw: string) => {
      if (readOnly || !enabled) return;
      let val = raw;

      if (singleLineMode) {
        val = val.replace(/\n|\r/g, "");
      }

      if (restriction && RESTRICTION_PATTERNS[restriction]) {
        const pattern = RESTRICTION_PATTERNS[restriction];
        val = val.split("").filter((c) => pattern.test(c)).join("");
      }

      if (maxLength !== undefined && val.length > maxLength) {
        val = val.slice(0, maxLength);
      }

      onChange(val || null);
    },
    [readOnly, enabled, singleLineMode, restriction, maxLength, onChange],
  );

  const positionViaPadding = useCallback(() => {
    const el = textareaRef.current;
    if (!el) return;

    el.style.paddingTop = "0px";
    el.style.paddingBottom = "0px";

    const clone = el.cloneNode(true) as HTMLTextAreaElement;
    clone.value = el.value;
    clone.style.position = "absolute";
    clone.style.left = "-9999px";
    clone.style.top = "0";
    clone.style.visibility = "hidden";
    clone.style.height = "auto";
    clone.style.maxHeight = "none";
    clone.style.padding = "0px";
    document.body.appendChild(clone);
    const contentHeight = clone.scrollHeight;
    document.body.removeChild(clone);

    const totalHeight = el.clientHeight;

    if (contentHeight >= totalHeight) return;

    const extra = totalHeight - contentHeight;
    if (vAlign === "middle") {
      el.style.paddingTop = Math.floor(extra / 2) + "px";
      el.style.paddingBottom = Math.ceil(extra / 2) + "px";
    } else if (vAlign === "bottom") {
      el.style.paddingTop = extra + "px";
    }
  }, [vAlign]);

  useLayoutEffect(() => {
    positionViaPadding();
  }, [positionViaPadding, displayValue]);

  const containerStyle: React.CSSProperties = {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    overflow: "hidden",
    border: production
      ? "2px solid rgba(234, 179, 8, 0.6)"
      : "1px solid rgba(59, 130, 246, 0.5)",
    borderRadius: "2px",
    background: config?.appearance?.backgroundColor ?? (production ? "rgba(254, 249, 195, 0.25)" : "rgba(255, 255, 255, 0.85)"),
    transition: "border-color 0.15s, box-shadow 0.15s",
    opacity: enabled ? 1 : 0.5,
    cursor: readOnly ? "default" : "text",
    ...(required && !readOnly ? { borderLeft: "3px solid #dc2626" } : {}),
  };

  const textareaStyle: React.CSSProperties = {
    width: "100%",
    height: "100%",
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
    lineHeight: LINE_HEIGHT_RATIO,
    whiteSpace: singleLineMode ? "nowrap" : "pre-wrap",
    overflow: "hidden",
    resize: "none",
  };

  return (
    <div style={containerStyle}>
      <textarea
        ref={textareaRef}
        value={displayValue}
        onChange={(e) => handleChange(e.target.value)}
        onBlur={onBlur}
        placeholder={placeholder ?? ""}
        maxLength={maxLength}
        readOnly={readOnly}
        disabled={!enabled}
        style={textareaStyle}
        className="runtime-field runtime-textarea-prototype"
      />
    </div>
  );
}

export const KeyboardTextFieldPrototype = memo(KeyboardTextFieldPrototypeInner);
