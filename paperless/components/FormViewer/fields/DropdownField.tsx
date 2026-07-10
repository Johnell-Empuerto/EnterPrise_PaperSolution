"use client";

import { useState } from "react";
import type { RuntimeField } from "@/types/runtime";

interface DropdownFieldProps {
  field: RuntimeField;
  onFocus?: () => void;
  style: React.CSSProperties;
}

/**
 * Select dropdown with yellow highlight.
 * Options can be specified via field metadata in the future.
 */
export function DropdownField({ field, onFocus, style }: DropdownFieldProps) {
  const [value, setValue] = useState(field.defaultValue ?? "");
  const [focused, setFocused] = useState(false);

  const selectStyle: React.CSSProperties = {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    backgroundColor: focused ? "#FFF9C4" : "#FFFDE7",
    border: focused ? "2px solid #FDD835" : "1px solid #F9A825",
    borderRadius: "2px",
    padding: "1px 4px",
    fontSize: `${Math.max(8, field.fontSize * 0.85)}px`,
    fontFamily: field.font ?? "inherit",
    fontWeight: field.bold ? "bold" : "normal",
    color: field.fontColor ?? "#333",
    outline: "none",
    boxShadow: focused ? "0 0 0 2px rgba(253, 216, 53, 0.3)" : "none",
    transition: "box-shadow 0.15s, background-color 0.15s",
    cursor: field.readOnly ? "default" : "pointer",
    opacity: field.readOnly ? 0.7 : 1,
  };

  return (
    <select
      style={{ ...style, ...selectStyle }}
      value={value}
      onChange={(e) => setValue(e.target.value)}
      onFocus={() => { setFocused(true); onFocus?.(); }}
      onBlur={() => setFocused(false)}
      disabled={field.readOnly}
      title={`${field.cellReference}${field.required ? " *" : ""}`}
    >
      <option value="" disabled>
        {field.placeholder ?? "Select..."}
      </option>
      <option value="Option 1">Option 1</option>
      <option value="Option 2">Option 2</option>
      <option value="Option 3">Option 3</option>
    </select>
  );
}
