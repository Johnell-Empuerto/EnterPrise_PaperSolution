"use client";

import type { ReactNode } from "react";

interface CanvasViewportProps {
  width: number;
  height: number;
  children: ReactNode;
}

export default function CanvasViewport({ width, height, children }: CanvasViewportProps) {
  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        overflow: "auto",
        position: "relative",
        background: "#e5e5e5",
      }}
    >
      <div
        style={{
          width,
          height,
          position: "relative",
          flexShrink: 0,
        }}
      >
        {children}
      </div>
    </div>
  );
}
