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
  /** Called when the field receives focus. */
  onFocus?: () => void;
}

/**
 * Dispatches to the correct field component based on the field's dataType.
 * Each field is absolutely positioned using the backend-provided pixel coordinates.
 *
 * Coordinate rule:
 *   left  = field.leftPx
 *   top   = field.topPx
 *   width = field.widthPx
 *   height = field.heightPx
 *
 * Positions are NEVER recalculated in the frontend.
 * The backend (CoordinateEngine) is the single source of truth.
 */
export function OverlayField({ field, onFocus }: OverlayFieldProps) {
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
        left: field.leftPx,
        top: field.topPx,
        width: field.widthPx,
        height: field.heightPx,
        zIndex: 10,
      }}
    >
      {renderField()}
    </div>
  );
}
