"use client";

import type { RuntimeField } from "@/types/runtime";
import { TextField } from "./fields/TextField";
import { NumberField } from "./fields/NumberField";
import { DateField } from "./fields/DateField";
import { CheckboxField } from "./fields/CheckboxField";
import { DropdownField } from "./fields/DropdownField";
import { SignatureField } from "./fields/SignatureField";

interface OverlayFieldProps {
  /** The runtime field definition from the backend. */
  field: RuntimeField;
  /** Page width in pixels for ratio-based rendering fallback. */
  pageWidthPx: number;
  /** Page height in pixels for ratio-based rendering fallback. */
  pageHeightPx: number;
  /** Called when the field receives focus. */
  onFocus?: () => void;
  /** Visual debug mode: renders red (computed) and green (actual cell) borders. */
  debug?: boolean;
}

/**
 * Dispatches to the correct field component based on the field's dataType.
 * Each field is absolutely positioned using backend-provided coordinates.
 *
 * Coordinate rule (priority order):
 *   1. Use leftPx/topPx/widthPx/heightPx if available (non-zero)
 *   2. Fall back to leftRatio*pageWidth / topRatio*pageHeight if ratios are available
 *   3. Default to 0,0,0,0 (invisible fallback)
 *
 * The backend (CoordinateEngine) is the single source of truth for positions.
 */
export function OverlayField({ field, pageWidthPx, pageHeightPx, onFocus, debug }: OverlayFieldProps) {
  // Resolve pixel coordinates with ratio fallback
  const leftPx = field.leftPx !== 0 ? field.leftPx : (field.leftRatio > 0 ? field.leftRatio * pageWidthPx : 0);
  const topPx = field.topPx !== 0 ? field.topPx : (field.topRatio > 0 ? field.topRatio * pageHeightPx : 0);
  const widthPx = field.widthPx !== 0 ? field.widthPx : (field.widthRatio > 0 ? field.widthRatio * pageWidthPx : 0);
  const heightPx = field.heightPx !== 0 ? field.heightPx : (field.heightRatio > 0 ? field.heightRatio * pageHeightPx : 0);

  // ── DEBUG: Log each field's resolved coordinates ──────────────────────
  console.log("[DEBUG] OverlayField:", {
    cellRef: field.cellReference,
    dataType: field.dataType,
    source: field.leftPx !== 0 ? "pixel" : field.leftRatio > 0 ? "ratio" : "zero",
    leftPx,
    topPx,
    widthPx,
    heightPx,
    raw_leftPx: field.leftPx,
    raw_topPx: field.topPx,
    raw_leftRatio: field.leftRatio,
    raw_topRatio: field.topRatio,
    pageWidthPx,
    pageHeightPx,
  });

  // Skips rendering if both pixel and ratio are zero (invalid field)
  if (leftPx === 0 && topPx === 0 && widthPx === 0 && heightPx === 0) {
    console.warn("[DEBUG] Skipping field — all dimensions are zero:", field.cellReference);
    return null;
  }

  // Fill style: the wrapper div provides the correct pixel dimensions via position:absolute.
  // Every child uses width:100%; height:100% to fill exactly those pixel boundaries.
  // boxSizing:border-box ensures padding/border are included in the 100% — no overflow.
  const fillStyle: React.CSSProperties = {
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
  };

  const renderField = () => {
    switch (field.dataType) {
      case "number":
        return <NumberField field={field} onFocus={onFocus} style={fillStyle} />;
      case "date":
        return <DateField field={field} onFocus={onFocus} style={fillStyle} />;
      case "checkbox":
        return <CheckboxField field={field} onFocus={onFocus} style={fillStyle} />;
      case "dropdown":
        return <DropdownField field={field} onFocus={onFocus} style={fillStyle} />;
      case "signature":
        return <SignatureField field={field} onFocus={onFocus} style={fillStyle} />;
      case "calculated":
        // Calculated fields are read-only by default
        return (
          <div
            style={fillStyle}
            className="bg-gray-50 border border-gray-200 rounded text-xs text-gray-500 flex items-center px-1.5 pointer-events-none"
            title={`${field.cellReference} (calculated)`}
          >
            {field.defaultValue ?? ""}
          </div>
        );
      case "text":
      default:
        return <TextField field={field} onFocus={onFocus} style={fillStyle} />;
    }
  };

  return (
    <div
      style={{
        position: "absolute",
        left: leftPx,
        top: topPx,
        width: widthPx,
        height: heightPx,
        zIndex: 10,
        ...(debug ? {
          outline: "2px solid #00FF00",
          outlineOffset: "-1px",
          backgroundColor: "rgba(0, 255, 0, 0.08)",
        } : {}),
      }}
    >
      {renderField()}
      {/* Debug label shows field dimensions */}
      {debug && (
        <div
          style={{
            position: "absolute",
            top: 0,
            left: 0,
            fontSize: "9px",
            lineHeight: 1,
            fontFamily: "monospace",
            color: "#00FF00",
            backgroundColor: "rgba(0,0,0,0.6)",
            padding: "1px 2px",
            borderRadius: "1px",
            pointerEvents: "none",
            whiteSpace: "nowrap",
            zIndex: 20,
          }}
        >
          {field.cellReference} {widthPx.toFixed(0)}x{heightPx.toFixed(0)}
        </div>
      )}
    </div>
  );
}
