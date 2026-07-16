import type { StateCreator } from "zustand";
import type { FieldErrors } from "@/types/runtime";

export interface ValidationSlice {
  errors: FieldErrors;
  warnings: Record<string, string[]>;

  setError: (fieldId: string, error: string | null) => void;
  setErrors: (errors: FieldErrors) => void;
  clearErrors: () => void;
  getError: (fieldId: string) => string | null;
  hasErrors: () => boolean;

  setWarning: (fieldId: string, warning: string) => void;
  clearWarnings: () => void;
  getWarnings: (fieldId: string) => string[];
}

export const createValidationSlice: StateCreator<ValidationSlice> = (set, get) => ({
  errors: {},
  warnings: {},

  setError: (fieldId: string, error: string | null) => {
    set((state) => ({
      errors: { ...state.errors, [fieldId]: error },
    }));
  },

  setErrors: (errors: FieldErrors) => set({ errors }),

  clearErrors: () => set({ errors: {} }),

  getError: (fieldId: string) => get().errors[fieldId] ?? null,

  hasErrors: () => {
    const errors = get().errors;
    return Object.values(errors).some((e) => e !== null && e !== undefined);
  },

  setWarning: (fieldId: string, warning: string) => {
    set((state) => {
      const existing = state.warnings[fieldId] ?? [];
      return {
        warnings: { ...state.warnings, [fieldId]: [...existing, warning] },
      };
    });
  },

  clearWarnings: () => set({ warnings: {} }),

  getWarnings: (fieldId: string) => get().warnings[fieldId] ?? [],
});
