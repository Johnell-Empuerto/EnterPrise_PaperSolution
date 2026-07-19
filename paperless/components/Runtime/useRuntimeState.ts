"use client";

import { useCallback } from "react";
import { useStore } from "zustand";
import { getDefaultStore } from "@/runtime/store";
import type { FieldValues, DirtyState } from "@/types/runtime";

export interface RuntimeState {
  values: FieldValues;
  dirty: DirtyState;
  lastUpdated: string | null;
  setValue: (overlayId: string, value: string | boolean | null) => void;
  reset: () => void;
  exportJson: () => string;
  isDirty: () => boolean;
  markAllClean: () => void;
}

/**
 * Runtime state management hook.
 *
 * BRIDGE IMPLEMENTATION (Phase 1 — Runtime V2):
 * Delegates to the Zustand store internally for backward compatibility.
 * Components that call this hook still work exactly as before.
 * New components should use `useRuntimeStore(selector)` directly
 * for per-field subscriptions and optimal re-renders.
 */
export function useRuntimeState(): RuntimeState {
  const store = getDefaultStore();

  const values = useStore(store, (s) => s.values);
  const dirty = useStore(store, (s) => s.dirty);

  const setValue = useCallback(
    (overlayId: string, value: string | boolean | null) => {
      store.getState().setValue(overlayId, value);
      store.getState().markDirty(overlayId);
    },
    [store],
  );

  const reset = useCallback(() => {
    store.getState().resetValues();
    store.getState().resetDirty();
  }, [store]);

  const exportJson = useCallback((): string => {
    return JSON.stringify(
      {
        values: store.getState().values,
        exportedAt: new Date().toISOString(),
        fieldCount: Object.keys(store.getState().values).length,
      },
      null,
      2,
    );
  }, [store]);

  const isDirty = useCallback((): boolean => {
    return store.getState().isDirty();
  }, [store]);

  const markAllClean = useCallback(() => {
    store.getState().markAllClean();
  }, [store]);

  return {
    values,
    dirty,
    lastUpdated: null,
    setValue,
    reset,
    exportJson,
    isDirty,
    markAllClean,
  };
}
