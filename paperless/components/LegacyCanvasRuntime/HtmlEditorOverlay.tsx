"use client";

import { useState, useRef, useEffect } from "react";
import type { LegacyRuntimeField } from "./LegacyCanvasRuntime";

interface HtmlEditorOverlayProps {
  field: LegacyRuntimeField;
  onSubmit: (value: string) => void;
  onCancel: () => void;
}

export default function HtmlEditorOverlay({
  field,
  onSubmit,
  onCancel,
}: HtmlEditorOverlayProps) {
  const [value, setValue] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      onSubmit(value);
    }
    if (e.key === "Escape") {
      onCancel();
    }
  };

  return (
    <input
      ref={inputRef}
      type="text"
      value={value}
      onChange={(e) => setValue(e.target.value)}
      onKeyDown={handleKeyDown}
      onBlur={() => onSubmit(value)}
      style={{
        position: "absolute",
        left: field.leftPx,
        top: field.topPx,
        width: field.widthPx,
        height: field.heightPx,
        margin: 0,
        padding: "1px 2px",
        border: "2px solid #00aa00",
        background: "rgba(255,255,255,0.95)",
        font: "11px Calibri, sans-serif",
        boxSizing: "border-box",
        outline: "none",
        zIndex: 10,
      }}
    />
  );
}
