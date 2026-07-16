import type { StateCreator } from "zustand";

export interface NavigationSlice {
  currentPageIndex: number;
  zoom: number;
  offsetX: number;
  offsetY: number;

  setPage: (index: number) => void;
  goNext: (pageCount: number) => void;
  goPrev: () => void;
  setZoom: (zoom: number) => void;
  setOffset: (x: number, y: number) => void;
  resetNavigation: () => void;
}

export const createNavigationSlice: StateCreator<NavigationSlice> = (set, get) => ({
  currentPageIndex: 0,
  zoom: 1,
  offsetX: 0,
  offsetY: 0,

  setPage: (index: number) => set({ currentPageIndex: index }),

  goNext: (pageCount: number) => {
    const next = get().currentPageIndex + 1;
    if (next < pageCount) {
      set({ currentPageIndex: next });
    }
  },

  goPrev: () => {
    const prev = get().currentPageIndex - 1;
    if (prev >= 0) {
      set({ currentPageIndex: prev });
    }
  },

  setZoom: (zoom: number) => {
    const clamped = Math.max(0.1, Math.min(5, zoom));
    set({ zoom: clamped });
  },

  setOffset: (x: number, y: number) => set({ offsetX: x, offsetY: y }),

  resetNavigation: () => set({ currentPageIndex: 0, zoom: 1, offsetX: 0, offsetY: 0 }),
});
