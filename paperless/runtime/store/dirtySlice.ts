import type { StateCreator } from "zustand";
import type { DirtyState } from "@/types/runtime";

export interface DirtySlice {
  dirty: DirtyState;

  markDirty: (fieldId: string) => void;
  markClean: (fieldId: string) => void;
  markAllClean: () => void;
  isDirty: () => boolean;
  dirtyCount: () => number;
  resetDirty: () => void;
}

export const createDirtySlice: StateCreator<DirtySlice> = (set, get) => ({
  dirty: {},

  markDirty: (fieldId: string) => {
    set((state) => ({
      dirty: { ...state.dirty, [fieldId]: true },
    }));
  },

  markClean: (fieldId: string) => {
    set((state) => {
      const next = { ...state.dirty };
      delete next[fieldId];
      return { dirty: next };
    });
  },

  markAllClean: () => set({ dirty: {} }),

  isDirty: () => Object.keys(get().dirty).length > 0,

  dirtyCount: () => Object.keys(get().dirty).length,

  resetDirty: () => set({ dirty: {} }),
});
