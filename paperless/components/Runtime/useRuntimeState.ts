"use client";

import { useState, useCallback, useRef } from "react";
import type { FieldValues, DirtyState } from "@/types/runtime";

export interface RuntimeState {
  values: FieldValues;
  dirty: DirtyState;
  lastUpdated: string | null;
  setValue: (overlayId: string, value: string | boolean | null) => void;
  reset: () => void;
  exportJson: () => string;
  isDirty: () => boolean;
}

/**
 * Runtime state management hook.
 * Stores form field values in React state only — no backend, no database.
 */
export function useRuntimeState(): RuntimeState {
  const [values, setValues] = useState<FieldValues>({});
  const [dirty, setDirty] = useState<DirtyState>({});
  const lastUpdatedRef = useRef<string | null>(null);

  const setValue = useCallback((overlayId: string, value: string | boolean | null) => {
    setValues((prev) => ({ ...prev, [overlayId]: value }));
    setDirty((prev) => ({ ...prev, [overlayId]: true }));
    lastUpdatedRef.current = new Date().toISOString();
  }, []);

  const reset = useCallback(() => {
    setValues({});
    setDirty({});
    lastUpdatedRef.current = null;
  }, []);

  const exportJson = useCallback((): string => {
    return JSON.stringify(
      {
        values,
        exportedAt: new Date().toISOString(),
        fieldCount: Object.keys(values).length,
      },
      null,
      2
    );
  }, [values]);

  const isDirty = useCallback((): boolean => {
    return Object.keys(dirty).length > 0;
  }, [dirty]);

  return {
    values,
    dirty,
    lastUpdated: lastUpdatedRef.current,
    setValue,
    reset,
    exportJson,
    isDirty,
  };
}
