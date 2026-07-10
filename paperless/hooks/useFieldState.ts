"use client";

import { useState, useCallback, useMemo } from "react";
import type { RuntimeField, FieldValues, FieldErrors, DirtyState } from "@/types/runtime";

interface UseFieldStateResult {
  /** Current values for all fields. */
  values: FieldValues;
  /** Validation errors keyed by field id. */
  errors: FieldErrors;
  /** Dirty state keyed by field id. */
  dirty: DirtyState;
  /** Whether any field is dirty. */
  isDirty: boolean;
  /** Whether all required fields have values. */
  isValid: boolean;
  /** Set a field value. */
  setValue: (fieldId: string, value: string | boolean | null) => void;
  /** Reset all fields to their default values. */
  reset: () => void;
  /** Validate a single field. */
  validateField: (field: RuntimeField) => string | null;
  /** Validate all fields. Returns true if all valid. */
  validateAll: (fields: RuntimeField[]) => boolean;
  /** Clear all errors. */
  clearErrors: () => void;
}

/**
 * Manages form field state: values, dirty tracking, validation.
 * Designed for the Yellow Editable Overlay (Phase 11J).
 */
export function useFieldState(fields: RuntimeField[]): UseFieldStateResult {
  // Initialize values from field defaults
  const initialValues = useMemo(() => {
    const vals: FieldValues = {};
    for (const f of fields) {
      if (f.dataType === "checkbox") {
        vals[f.id] = f.defaultValue === "true" || f.defaultValue === "yes";
      } else {
        vals[f.id] = f.defaultValue ?? null;
      }
    }
    return vals;
  }, [fields]);

  const [values, setValues] = useState<FieldValues>(initialValues);
  const [errors, setErrors] = useState<FieldErrors>({});
  const [dirty, setDirty] = useState<DirtyState>({});

  const isDirty = useMemo(
    () => Object.values(dirty).some((d) => d),
    [dirty]
  );

  const isValid = useMemo(() => {
    for (const field of fields) {
      if (field.required) {
        const val = values[field.id];
        if (val === null || val === undefined || val === "") {
          return false;
        }
      }
    }
    return Object.keys(errors).length === 0 ||
      Object.values(errors).every((e) => e === null);
  }, [fields, values, errors]);

  const setValue = useCallback(
    (fieldId: string, value: string | boolean | null) => {
      setValues((prev) => ({ ...prev, [fieldId]: value }));
      setDirty((prev) => ({ ...prev, [fieldId]: true }));
      // Clear error for this field on change
      setErrors((prev) => ({ ...prev, [fieldId]: null }));
    },
    []
  );

  const reset = useCallback(() => {
    setValues(initialValues);
    setErrors({});
    setDirty({});
  }, [initialValues]);

  const validateField = useCallback(
    (field: RuntimeField): string | null => {
      const value = values[field.id];

      // Required check
      if (field.required) {
        if (value === null || value === undefined || value === "") {
          const msg = field.validationMessage ?? `${field.cellReference} is required`;
          setErrors((prev) => ({ ...prev, [field.id]: msg }));
          return msg;
        }
      }

      // Validation pattern (regex)
      if (field.validationPattern && typeof value === "string" && value.length > 0) {
        try {
          const regex = new RegExp(field.validationPattern);
          if (!regex.test(value)) {
            const msg = field.validationMessage ?? `Invalid format for ${field.cellReference}`;
            setErrors((prev) => ({ ...prev, [field.id]: msg }));
            return msg;
          }
        } catch {
          // Invalid regex in field definition — skip
        }
      }

      // Max length
      if (field.maxLength > 0 && typeof value === "string" && value.length > field.maxLength) {
        const msg = `Maximum ${field.maxLength} characters`;
        setErrors((prev) => ({ ...prev, [field.id]: msg }));
        return msg;
      }

      // Number validation
      if (field.dataType === "number" && typeof value === "string" && value.length > 0) {
        if (isNaN(Number(value))) {
          const msg = `Must be a valid number`;
          setErrors((prev) => ({ ...prev, [field.id]: msg }));
          return msg;
        }
      }

      // Date validation
      if (field.dataType === "date" && typeof value === "string" && value.length > 0) {
        if (isNaN(Date.parse(value))) {
          const msg = `Must be a valid date`;
          setErrors((prev) => ({ ...prev, [field.id]: msg }));
          return msg;
        }
      }

      // Clear error
      setErrors((prev) => ({ ...prev, [field.id]: null }));
      return null;
    },
    [values]
  );

  const validateAll = useCallback(
    (allFields: RuntimeField[]): boolean => {
      let allValid = true;
      const newErrors: FieldErrors = {};

      for (const field of allFields) {
        const value = values[field.id];
        let error: string | null = null;

        if (field.required && (value === null || value === undefined || value === "")) {
          error = field.validationMessage ?? `${field.cellReference} is required`;
          allValid = false;
        }

        if (!error && field.validationPattern && typeof value === "string" && value.length > 0) {
          try {
            const regex = new RegExp(field.validationPattern);
            if (!regex.test(value)) {
              error = field.validationMessage ?? `Invalid format`;
              allValid = false;
            }
          } catch { /* skip */ }
        }

        if (!error && field.maxLength > 0 && typeof value === "string" && value.length > field.maxLength) {
          error = `Maximum ${field.maxLength} characters`;
          allValid = false;
        }

        if (!error && field.dataType === "number" && typeof value === "string" && value.length > 0 && isNaN(Number(value))) {
          error = `Must be a valid number`;
          allValid = false;
        }

        if (!error && field.dataType === "date" && typeof value === "string" && value.length > 0 && isNaN(Date.parse(value))) {
          error = `Must be a valid date`;
          allValid = false;
        }

        newErrors[field.id] = error;
      }

      setErrors(newErrors);
      return allValid;
    },
    [values]
  );

  const clearErrors = useCallback(() => {
    setErrors({});
  }, []);

  return {
    values,
    errors,
    dirty,
    isDirty,
    isValid,
    setValue,
    reset,
    validateField,
    validateAll,
    clearErrors,
  };
}
