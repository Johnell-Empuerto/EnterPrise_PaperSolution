"use client";

import type { ReactNode } from "react";

/**
 * PageSurface — the browser equivalent of the WPF Canvas.
 *
 * This is the SINGLE coordinate system for the entire form page.
 * Every rendering layer (background, fields, annotations, etc.)
 * shares this precise coordinate space.
 *
 * Properties:
 * - Fixed pixel dimensions matching the PNG/overlay coordinate space
 * - position: relative — creates positioning context for absolute children
 * - overflow: hidden — content is clipped to the page surface
 * - No responsive sizing, no flex, no percentage widths
 * - flexShrink: 0 — prevents the surface from being shrunk by flex parents
 *
 * Legacy equivalent: WPF Canvas with Width/Height set to bitmap pixel size.
 */
export interface PageSurfaceProps {
  /** Width in CSS pixels — must match the background image pixel width */
  widthPx: number;
  /** Height in CSS pixels — must match the background image pixel height */
  heightPx: number;
  /** Child layers: BackgroundLayer, FieldLayer, AnnotationLayer, etc. */
  children: ReactNode;
}

export function PageSurface({ widthPx, heightPx, children }: PageSurfaceProps) {
  return (
    <div
      data-page-surface
      style={{
        position: "relative",
        width: widthPx,
        height: heightPx,
        overflow: "hidden",
        flexShrink: 0,
        /* No flex, no margins, no responsive sizing — pure fixed canvas */
      }}
    >
      {children}
    </div>
  );
}
