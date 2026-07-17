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
import { shapeText, LINE_HEIGHT_RATIO, MIN_FONT_SIZE_PT, MEASURE_TOLERANCE_PX } from "@/runtime/config/fontShaper";

const RESTRICTION_PATTERNS: Record<string, RegExp> = {
  letters: /^[A-Za-z]*$/,
  numbers: /^[0-9]*$/,
  alphanumeric: /^[A-Za-z0-9]*$/,
};

function KeyboardTextFieldInner({
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

  const relayout = useCallback(() => {
    const el = textareaRef.current;
    if (!el) return;

    el.style.fontSize = `${fontSize}pt`;
    el.style.paddingTop = "1px";
    el.style.paddingBottom = "1px";

    const totalHeight = el.clientHeight;
    const basePad = 1;
    const contentAreaHeight = totalHeight - basePad * 2;
    const availWidth = el.clientWidth - 6 - MEASURE_TOLERANCE_PX;

    if (availWidth <= 0 || contentAreaHeight <= 0) return;

    const result = shapeText(displayValue, {
      fontFamily,
      fontWeight,
      originalFontSizePt: fontSize,
      minFontSizePt: MIN_FONT_SIZE_PT,
      lineHeightRatio: LINE_HEIGHT_RATIO,
      availableWidthPx: availWidth,
      availableHeightPx: contentAreaHeight,
    });

    el.style.fontSize = `${result.fontSize}pt`;

    const extra = contentAreaHeight - result.contentHeightPx;
    if (extra > 0 && vAlign === "middle") {
      el.style.paddingTop = (basePad + Math.floor(extra / 2)) + "px";
      el.style.paddingBottom = (basePad + Math.ceil(extra / 2)) + "px";
    } else if (extra > 0 && vAlign === "bottom") {
      el.style.paddingTop = (basePad + extra) + "px";
      el.style.paddingBottom = basePad + "px";
    } else {
      el.style.paddingTop = basePad + "px";
      el.style.paddingBottom = basePad + "px";
    }
  }, [displayValue, fontFamily, fontWeight, fontSize, vAlign]);

  useLayoutEffect(() => {
    relayout();
  }, [relayout]);

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
    whiteSpace: "pre-wrap",
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
        placeholder={kt.placeholder || (placeholder ?? "")}
        maxLength={maxLength}
        readOnly={readOnly}
        disabled={!enabled}
        style={textareaStyle}
        className="runtime-field runtime-textarea"
      />
    </div>
  );
}

export const KeyboardTextField = memo(KeyboardTextFieldInner);
