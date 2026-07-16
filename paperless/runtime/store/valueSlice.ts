import type { StateCreator } from "zustand";
import type { FieldValues } from "@/types/runtime";

export interface ValueSlice {
  values: FieldValues;

  setValue: (fieldId: string, value: string | boolean | null) => void;
  setMultipleValues: (values: FieldValues) => void;
  getValue: (fieldId: string) => string | boolean | null;
  resetValues: () => void;
}

export const createValueSlice: StateCreator<ValueSlice> = (set, get) => ({
  values: {},

  setValue: (fieldId: string, value: string | boolean | null) => {
    set((state) => ({
      values: { ...state.values, [fieldId]: value },
    }));
  },

  setMultipleValues: (values: FieldValues) => {
    set((state) => ({
      values: { ...state.values, ...values },
    }));
  },

  getValue: (fieldId: string) => {
    return get().values[fieldId] ?? null;
  },

  resetValues: () => set({ values: {} }),
});
