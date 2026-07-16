import type { StateCreator } from "zustand";
import type { RuntimeForm, RuntimeSheet } from "@/types/runtime";

export interface FormSlice {
  formId: string | null;
  title: string;
  pages: RuntimeSheet[];
  currentPageIndex: number;
  isLoading: boolean;
  error: string | null;

  loadForm: (form: RuntimeForm) => void;
  setCurrentPage: (index: number) => void;
  setLoading: (loading: boolean) => void;
  setFormError: (error: string | null) => void;
  getCurrentPage: () => RuntimeSheet | null;
  getPageCount: () => number;
  resetForm: () => void;
}

export const createFormSlice: StateCreator<FormSlice> = (set, get) => ({
  formId: null,
  title: "",
  pages: [],
  currentPageIndex: 0,
  isLoading: false,
  error: null,

  loadForm: (form: RuntimeForm) => {
    set({
      formId: form.title,
      title: form.workbookName,
      pages: form.sheets,
      currentPageIndex: 0,
      isLoading: false,
      error: null,
    });
  },

  setCurrentPage: (index: number) => {
    const pageCount = get().pages.length;
    if (index >= 0 && index < pageCount) {
      set({ currentPageIndex: index });
    }
  },

  setLoading: (isLoading: boolean) => set({ isLoading }),

  setFormError: (error: string | null) => set({ error, isLoading: false }),

  getCurrentPage: () => {
    const { pages, currentPageIndex } = get();
    return pages[currentPageIndex] ?? null;
  },

  getPageCount: () => get().pages.length,

  resetForm: () => {
    set({
      formId: null,
      title: "",
      pages: [],
      currentPageIndex: 0,
      isLoading: false,
      error: null,
    });
  },
});
