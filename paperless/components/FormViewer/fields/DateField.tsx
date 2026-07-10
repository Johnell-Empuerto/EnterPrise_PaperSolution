"use client";

import { useState } from "react";
import type { RuntimeField } from "@/types/runtime";

interface DateFieldProps {
  field: RuntimeField;
  onFocus?: () => void;
  style: React.CSSProperties;
}

/**
 * Date input field with yellow highlight.
 * Uses browser-native date picker.
 */
export function DateField({ field, onFocus, style }: DateFieldProps) {
  const [value, setValue] = useState(field.defaultValue ?? "");
  const [focused, setFocused] = useState(false);

  const inputStyle: React.CSSProperties = {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    backgroundColor: focused ? "#FFF9C4" : "#FFFDE7",
    border: focused ? "2px solid #FDD835" : "1px solid #F9A825",
    borderRadius: "2px",
    padding: "1px 4px",
    fontSize: `${Math.max(8, field.fontSize * 0.85)}px`,
    fontFamily: field.font ?? "inherit",
    color: field.fontColor ?? "#333",
    outline: "none",
    boxShadow: focused ? "0 0 0 2px rgba(253, 216, 53, 0.3)" : "none",
    transition: "box-shadow 0.15s, background-color 0.15s",
    cursor: field.readOnly ? "default" : "text",
    opacity: field.readOnly ? 0.7 : 1,
  };

  return (
    <input
      type="date"
      style={{ ...style, ...inputStyle }}
      value={value}
      onChange={(e) => setValue(e.target.value)}
      onFocus={() => { setFocused(true); onFocus?.(); }}
      onBlur={() => setFocused(false)}
      readOnly={field.readOnly}
      title={`${field.cellReference}${field.required ? " *" : ""}`}
    />
  );
}
