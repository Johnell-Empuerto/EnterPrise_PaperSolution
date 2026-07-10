"use client";

import { useState, useRef } from "react";
import type { RuntimeField } from "@/types/runtime";

interface SignatureFieldProps {
  field: RuntimeField;
  onFocus?: () => void;
  style: React.CSSProperties;
}

/**
 * Signature field with yellow highlight.
 * Clicking the field toggles a "signed" state with a visual signature line.
 * Future enhancement: actual canvas-based signature capture.
 */
export function SignatureField({ field, onFocus, style }: SignatureFieldProps) {
  const [signed, setSigned] = useState(false);
  const [focused, setFocused] = useState(false);
  const canvasRef = useRef<HTMLCanvasElement>(null);

  const handleClick = () => {
    if (field.readOnly) return;
    if (!signed) {
      setSigned(true);
      onFocus?.();
      // Draw a simple signature line on the canvas
      const canvas = canvasRef.current;
      if (canvas) {
        const ctx = canvas.getContext("2d");
        if (ctx) {
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          ctx.strokeStyle = "#333";
          ctx.lineWidth = 1.5;
          ctx.beginPath();
          ctx.moveTo(canvas.width * 0.1, canvas.height * 0.7);
          ctx.bezierCurveTo(
            canvas.width * 0.3, canvas.height * 0.3,
            canvas.width * 0.4, canvas.height * 0.9,
            canvas.width * 0.5, canvas.height * 0.5
          );
          ctx.bezierCurveTo(
            canvas.width * 0.6, canvas.height * 0.1,
            canvas.width * 0.7, canvas.height * 0.6,
            canvas.width * 0.9, canvas.height * 0.3
          );
          ctx.stroke();
        }
      }
    }
  };

  const containerStyle: React.CSSProperties = {
    ...style,
    boxSizing: "border-box",
    backgroundColor: focused ? "#FFF9C4" : "#FFFDE7",
    border: "none",
    borderRadius: "2px",
    cursor: field.readOnly ? "default" : "pointer",
    boxShadow: focused
      ? "inset 0 0 0 2px #FDD835, 0 0 0 2px rgba(253, 216, 53, 0.3)"
      : "inset 0 0 0 1px #F9A825",
    transition: "box-shadow 0.15s, background-color 0.15s",
    overflow: "hidden",
    opacity: field.readOnly ? 0.7 : 1,
  };

  return (
    <div
      style={containerStyle}
      onClick={handleClick}
      onMouseEnter={() => setFocused(true)}
      onMouseLeave={() => setFocused(false)}
      title={`${field.cellReference}${field.required ? " *" : ""}`}
    >
      <canvas
        ref={canvasRef}
        width={field.widthPx}
        height={field.heightPx}
        style={{ display: "block", width: "100%", height: "100%" }}
      />
      {!signed && (
        <div
          style={{
            position: "absolute",
            inset: 0,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: `${Math.max(8, field.fontSize * 0.75)}px`,
            color: "#999",
            pointerEvents: "none",
          }}
        >
          Click to sign
        </div>
      )}
    </div>
  );
}
