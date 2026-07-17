"use client";

import { memo } from "react";
import type { OverlayModel } from "@/types/overlay";
import { getFieldComponent } from "@/runtime/components/fields";
import { useRuntimeStore } from "@/runtime/store";
import { ValidationService } from "@/runtime/services/ValidationService";

export interface RuntimeFieldProps {
  overlay: OverlayModel;
  value: string | boolean | null;
  onChange: (value: string | boolean | null) => void;
  onSelect?: (fieldId: string) => void;
  production?: boolean;
  usePixelUnits?: boolean;
}

function RuntimeFieldInner({ overlay, production, usePixelUnits, onSelect }: RuntimeFieldProps) {
  const cfg = overlay.config;
  const visible = cfg?.behavior?.visible ?? true;
  const readOnly = cfg?.behavior?.readOnly ?? false;

  // Per-field selectors — only subscribe to THIS field's value and error
  const fieldId = overlay.id;
  const value = useRuntimeStore((s) => s.values[fieldId]);
  const error = useRuntimeStore((s) => s.errors[fieldId]);
  const setValue = useRuntimeStore((s) => s.setValue);
  const markDirty = useRuntimeStore((s) => s.markDirty);
  const setError = useRuntimeStore((s) => s.setError);

  // Look up the field component from the registry (not a switch-case)
  const FieldComponent = getFieldComponent(overlay.type);
  if (!FieldComponent) return null;

  const unit = usePixelUnits ? "px" : "pt";
  const fieldStyle: React.CSSProperties = {
    position: "absolute",
    left: `${overlay.leftPt}${unit}`,
    top: `${overlay.topPt}${unit}`,
    width: `${overlay.widthPt}${unit}`,
    height: `${overlay.heightPt}${unit}`,
    display: visible ? undefined : "none",
  };

  const handleChange = (newValue: string | boolean | null) => {
    setValue(fieldId, newValue);
    if (!readOnly) {
      markDirty(fieldId);
    }
  };

  // Validate on blur — uses pure ValidationService, no store logic here
  const handleBlur = () => {
    if (overlay.metadata?.dataType) {
      const fieldDef = {
        id: fieldId,
        cellReference: overlay.cell,
        dataType: overlay.metadata.dataType as any,
        required: cfg?.behavior?.required ?? false,
        maxLength: cfg?.input?.maxLength ?? 0,
        validationPattern: null,
        validationMessage: null,
      };
      const err = ValidationService.validateField(fieldDef, value);
      setError(fieldId, err);
    }
  };

  const handleClick = () => {
    onSelect?.(fieldId);
  };

  return (
    <div
      style={fieldStyle}
      data-field-id={fieldId}
      onClick={handleClick}
    >
      {cfg?.behavior?.required && (
        <span
          style={{
            position: "absolute",
            top: 0,
            right: 0,
            transform: "translate(50%, -50%)",
            fontSize: "10px",
            color: "#dc2626",
            fontWeight: "bold",
            zIndex: 1,
            pointerEvents: "none",
          }}
        >
          *
        </span>
      )}
      <FieldComponent
        overlay={overlay}
        value={value}
        onChange={handleChange}
        onBlur={handleBlur}
        production={production}
        config={cfg}
        readOnly={readOnly}
        required={cfg?.behavior?.required}
        placeholder={overlay.metadata?.placeholder as string | undefined}
      />
    </div>
  );
}

export const RuntimeField = memo(RuntimeFieldInner);
