import { create, useStore } from "zustand";
import { shallow } from "zustand/shallow";
import type { RuntimeForm } from "@/types/runtime";
import { createFormSlice, type FormSlice } from "./formSlice";
import { createValueSlice, type ValueSlice } from "./valueSlice";
import { createDirtySlice, type DirtySlice } from "./dirtySlice";
import { createValidationSlice, type ValidationSlice } from "./validationSlice";
import { createNavigationSlice, type NavigationSlice } from "./navigationSlice";

export type RuntimeStore = FormSlice & ValueSlice & DirtySlice & ValidationSlice & NavigationSlice;

/**
 * Creates a new RuntimeStore instance.
 * Multiple stores can exist per form session.
 */
export function createRuntimeStore() {
  return create<RuntimeStore>()((...a) => ({
    ...createFormSlice(...a),
    ...createValueSlice(...a),
    ...createDirtySlice(...a),
    ...createValidationSlice(...a),
    ...createNavigationSlice(...a),
  }));
}

let defaultStore: ReturnType<typeof createRuntimeStore> | null = null;

/**
 * Returns the default runtime store, creating it if necessary.
 * In production, a single store instance is shared by all components.
 */
export function getDefaultStore() {
  if (!defaultStore) {
    defaultStore = createRuntimeStore();
  }
  return defaultStore;
}

/**
 * Replaces the default store instance.
 * Used during testing or when resetting state.
 */
export function setDefaultStore(store: ReturnType<typeof createRuntimeStore>) {
  defaultStore = store;
}

/**
 * Hook to subscribe to the default store with a selector.
 * Components pick which slice/field to subscribe to for optimal re-renders.
 */
export function useRuntimeStore<T>(selector: (state: RuntimeStore) => T): T {
  const store = getDefaultStore();
  return useStore(store, selector);
}

/**
 * Initializes the store with form data.
 * Sets form metadata and initializes field values from defaults.
 */
export function initializeStore(
  store: ReturnType<typeof createRuntimeStore>,
  form: RuntimeForm
) {
  const state = store.getState();

  state.loadForm(form);

  const initialValues: Record<string, string | boolean | null> = {};
  for (const sheet of form.sheets) {
    for (const field of sheet.fields) {
      if (field.dataType === "checkbox") {
        initialValues[field.id] = field.defaultValue === "true" || field.defaultValue === "yes";
      } else {
        initialValues[field.id] = field.defaultValue ?? null;
      }
    }
  }
  state.setMultipleValues(initialValues);
}

export type RuntimeStoreApi = ReturnType<typeof createRuntimeStore>;
