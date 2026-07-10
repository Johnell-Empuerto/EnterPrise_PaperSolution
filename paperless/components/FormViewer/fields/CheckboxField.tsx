"use client";

import { useState } from "react";
import type { RuntimeField } from "@/types/runtime";

interface CheckboxFieldProps {
  field: RuntimeField;
  onFocus?: () => void;
  style: React.CSSProperties;
}

/**
 * Checkbox input with yellow highlight.
 * Centered within the field bounding box from the backend.
 */
export function CheckboxField({ field, onFocus, style }: CheckboxFieldProps) {
  const [checked, setChecked] = useState(
    field.defaultValue === "true" || field.defaultValue === "yes" || field.defaultValue === "☑"
  );
  const [focused, setFocused] = useState(false);

  // Compute checkbox size (fit within the field, max 20px)
  const boxSize = Math.min(field.widthPx, field.heightPx, 20);
  const offsetX = (field.widthPx - boxSize) / 2;
  const offsetY = (field.heightPx - boxSize) / 2;

  return (
    <div
      style={{
        ...style,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: field.readOnly ? "default" : "pointer",
        opacity: field.readOnly ? 0.7 : 1,
      }}
      onClick={() => {
        if (!field.readOnly) {
          setChecked((c) => !c);
          onFocus?.();
        }
      }}
      title={`${field.cellReference}${field.required ? " *" : ""}`}
    >
      <div
        style={{
          width: boxSize,
          height: boxSize,
          backgroundColor: focused ? "#FFF9C4" : "#FFFDE7",
          border: focused ? "2px solid #FDD835" : "1px solid #F9A825",
          borderRadius: "3px",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          boxShadow: focused ? "0 0 0 2px rgba(253, 216, 53, 0.3)" : "none",
          transition: "box-shadow 0.15s, background-color 0.15s",
          fontSize: `${Math.max(10, boxSize - 4)}px`,
          lineHeight: 1,
          userSelect: "none",
        }}
        onMouseEnter={() => setFocused(true)}
        onMouseLeave={() => setFocused(false)}
      >
        {checked ? (
          <span style={{ color: "#333" }}>✓</span>
        ) : null}
      </div>
    </div>
  );
}
