"use client";

import type { OverlayModel } from "@/types/overlay";
import { TextField } from "./fields/TextField";
import { CheckboxField } from "./fields/CheckboxField";
import { DateField } from "./fields/DateField";
import { NumberField } from "./fields/NumberField";
import { SignatureField } from "./fields/SignatureField";

export interface RuntimeFieldProps {
  overlay: OverlayModel;
  value: string | boolean | null;
  onChange: (value: string | boolean | null) => void;
  /** Production mode: use yellow styling for all fields */
  production?: boolean;
  /** Use px units instead of pt (for COM-based coordinate system where pixels are source of truth) */
  usePixelUnits?: boolean;
}

/**
 * Dispatches to the correct field component based on overlay type.
 * Every field uses OverlayModel coordinates — no coordinate calculations inside components.
 */
export function RuntimeField({ overlay, value, onChange, production, usePixelUnits }: RuntimeFieldProps) {
  // Container div positioned at the overlay coordinates
  const unit = usePixelUnits ? "px" : "pt";
  const fieldStyle: React.CSSProperties = {
    position: "absolute",
    left: `${overlay.leftPt}${unit}`,
    top: `${overlay.topPt}${unit}`,
    width: `${overlay.widthPt}${unit}`,
    height: `${overlay.heightPt}${unit}`,
    pointerEvents: "auto", // Inputs must be interactive
    zIndex: 26,
  };

  const commonProps = {
    overlay,
    value,
    onChange,
    production,
  };

  let field: React.ReactNode;

  switch (overlay.type) {
    case "textbox":
      field = <TextField {...commonProps} />;
      break;
    case "checkbox":
      field = <CheckboxField {...commonProps} />;
      break;
    case "date":
      field = <DateField {...commonProps} />;
      break;
    case "number":
      field = <NumberField {...commonProps} />;
      break;
    case "signature":
      field = <SignatureField {...commonProps} />;
      break;
    default:
      return null;
  }

  return (
    <div
      style={fieldStyle}
      onMouseDown={(e) => e.stopPropagation()}
    >
      {field}
    </div>
  );
}
