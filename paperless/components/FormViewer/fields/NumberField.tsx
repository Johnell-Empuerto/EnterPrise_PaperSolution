"use client";

import { useState } from "react";
import type { RuntimeField } from "@/types/runtime";

interface NumberFieldProps {
  field: RuntimeField;
  onFocus?: () => void;
  style: React.CSSProperties;
}

/**
 * Number input field with yellow highlight.
 * Allows decimal and negative values.
 */
export function NumberField({ field, onFocus, style }: NumberFieldProps) {
  const [value, setValue] = useState(field.defaultValue ?? "");
  const [focused, setFocused] = useState(false);

  const inputStyle: React.CSSProperties = {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    lineHeight: 1.2,
    backgroundColor: focused ? "#FFF9C4" : "#FFFDE7",
    border: "none",
    borderRadius: "2px",
    padding: "1px 4px",
    fontSize: `${Math.max(8, field.fontSize * 0.85)}px`,
    fontFamily: field.font ?? "inherit",
    fontWeight: field.bold ? "bold" : "normal",
    color: field.fontColor ?? "#333",
    textAlign: (field.alignment === "right" || field.alignment === "general") ? "right" : "left",
    outline: "none",
    boxShadow: focused
      ? "inset 0 0 0 2px #FDD835, 0 0 0 2px rgba(253, 216, 53, 0.3)"
      : "inset 0 0 0 1px #F9A825",
    transition: "box-shadow 0.15s, background-color 0.15s",
    cursor: field.readOnly ? "default" : "text",
    opacity: field.readOnly ? 0.7 : 1,
  };

  return (
    <input
      type="number"
      style={{ ...style, ...inputStyle }}
      value={value}
      onChange={(e) => setValue(e.target.value)}
      onFocus={() => { setFocused(true); onFocus?.(); }}
      onBlur={() => setFocused(false)}
      placeholder={field.placeholder ?? field.cellReference}
      readOnly={field.readOnly}
      step="any"
      title={`${field.cellReference}${field.required ? " *" : ""}`}
    />
  );
}
