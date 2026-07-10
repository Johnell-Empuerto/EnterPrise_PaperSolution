"use client";

import { useState } from "react";
import type { RuntimeField } from "@/types/runtime";

interface TextFieldProps {
  field: RuntimeField;
  onFocus?: () => void;
  style: React.CSSProperties;
}

/**
 * Text input field with yellow highlight (editable overlay).
 * Multiline when the field height suggests more than one line of text.
 */
export function TextField({ field, onFocus, style }: TextFieldProps) {
  const [value, setValue] = useState(field.defaultValue ?? "");
  const [focused, setFocused] = useState(false);

  // Determine if textarea is more appropriate based on height
  const isMultiline = field.heightPx > 28;

  const inputStyle: React.CSSProperties = {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    lineHeight: 1.2,
    backgroundColor: focused ? "#FFF9C4" : "#FFFDE7",
    border: "none",
    borderRadius: "2px",
    padding: isMultiline ? "2px 4px" : "1px 4px",
    fontSize: `${Math.max(8, field.fontSize * 0.85)}px`,
    fontFamily: field.font ?? "inherit",
    fontWeight: field.bold ? "bold" : "normal",
    color: field.fontColor ?? "#333",
    outline: "none",
    resize: "none",
    overflow: isMultiline ? "auto" : "hidden",
    boxShadow: focused
      ? "inset 0 0 0 2px #FDD835, 0 0 0 2px rgba(253, 216, 53, 0.3)"
      : "inset 0 0 0 1px #F9A825",
    transition: "box-shadow 0.15s, background-color 0.15s",
    cursor: field.readOnly ? "default" : "text",
    opacity: field.readOnly ? 0.7 : 1,
  };

  if (isMultiline) {
    const lines = Math.max(2, Math.round(field.heightPx / 18));
    return (
      <textarea
        style={{ ...style, ...inputStyle }}
        value={value}
        onChange={(e) => setValue(e.target.value)}
        onFocus={() => { setFocused(true); onFocus?.(); }}
        onBlur={() => setFocused(false)}
        placeholder={field.placeholder ?? field.cellReference}
        readOnly={field.readOnly}
        maxLength={field.maxLength > 0 ? field.maxLength : undefined}
        rows={lines}
        title={`${field.cellReference}${field.required ? " *" : ""}`}
      />
    );
  }

  return (
    <input
      type="text"
      style={{ ...style, ...inputStyle }}
      value={value}
      onChange={(e) => setValue(e.target.value)}
      onFocus={() => { setFocused(true); onFocus?.(); }}
      onBlur={() => setFocused(false)}
      placeholder={field.placeholder ?? field.cellReference}
      readOnly={field.readOnly}
      maxLength={field.maxLength > 0 ? field.maxLength : undefined}
      title={`${field.cellReference}${field.required ? " *" : ""}`}
    />
  );
}
