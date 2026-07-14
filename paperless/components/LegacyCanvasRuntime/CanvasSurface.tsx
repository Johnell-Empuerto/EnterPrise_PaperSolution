"use client";

import { useRef, useEffect, useCallback } from "react";
import type { LegacyRuntimeField, DebugFlags } from "./LegacyCanvasRuntime";
import { renderFrame } from "./CanvasRenderer";

interface CanvasSurfaceProps {
  width: number;
  height: number;
  backgroundImage: HTMLImageElement | null;
  fields: LegacyRuntimeField[];
  editingFieldId: string | null;
  debug: DebugFlags;
  onClick: (pageX: number, pageY: number) => void;
}

export default function CanvasSurface({
  width,
  height,
  backgroundImage,
  fields,
  editingFieldId,
  debug,
  onClick,
}: CanvasSurfaceProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const hasRenderedRef = useRef(false);

  useEffect(() => {
    // Stage 7: CanvasSurface useEffect
    console.log("[Stage 7] CanvasSurface useEffect triggered");
    console.log("[Stage 7] Props:", {
      width,
      height,
      hasBackgroundImage: backgroundImage !== null,
      fieldCount: fields.length,
      editingFieldId,
    });

    const canvas = canvasRef.current;
    if (!canvas) {
      console.error("[Stage 7] FATAL: canvasRef.current is null");
      return;
    }

    console.log("[Stage 7] Canvas element found. HTML attributes:", {
      width: canvas.width,
      height: canvas.height,
    });

    if (canvas.width !== width || canvas.height !== height) {
      console.warn("[Stage 7] Canvas dimensions mismatch! HTML:", canvas.width, "x", canvas.height, "Expected:", width, "x", height);
    }

    const ctx = canvas.getContext("2d");
    if (!ctx) {
      console.error("[Stage 7] FATAL: getContext('2d') returned null");
      return;
    }

    console.log("[Stage 7] Canvas 2D context obtained");

    // Stage 8: Call renderFrame (logging inside)
    renderFrame(ctx, backgroundImage, fields, editingFieldId, debug, width, height);

    // Stage 10: Post-render verification via requestAnimationFrame
    // This runs after the browser has painted the frame
    requestAnimationFrame(() => {
      console.log("[Stage 10] requestAnimationFrame callback fired");

      if (!canvas) return;
      const checkCtx = canvas.getContext("2d");
      if (!checkCtx) return;

      // Stage 11: Check first pixels
      try {
        const pixelData = checkCtx.getImageData(0, 0, 10, 10);
        let hasNonWhitePixel = false;
        let hasRedPixel = false;
        for (let i = 0; i < pixelData.data.length; i += 4) {
          const r = pixelData.data[i];
          const g = pixelData.data[i + 1];
          const b = pixelData.data[i + 2];
          const a = pixelData.data[i + 3];
          if (r !== 255 || g !== 255 || b !== 255) {
            hasNonWhitePixel = true;
          }
          if (r > 200 && g < 80 && b < 80) {
            hasRedPixel = true;
          }
        }
        console.log("[Stage 11] Canvas pixel check (10x10 top-left):");
        console.log("[Stage 11]   Has non-white pixel:", hasNonWhitePixel);
        console.log("[Stage 11]   Has red pixel:", hasRedPixel);
        console.log("[Stage 11]   First pixel RGBA:", pixelData.data[0], pixelData.data[1], pixelData.data[2], pixelData.data[3]);
        console.log("[Stage 11]   Pixels 0-3:",
          `[${pixelData.data[0]},${pixelData.data[1]},${pixelData.data[2]},${pixelData.data[3]}]`,
          `[${pixelData.data[4]},${pixelData.data[5]},${pixelData.data[6]},${pixelData.data[7]}]`,
          `[${pixelData.data[8]},${pixelData.data[9]},${pixelData.data[10]},${pixelData.data[11]}]`
        );

        if (!hasRedPixel && hasNonWhitePixel) {
          console.log("[Stage 11] Background IS rendering (non-white pixels found) but red diagnostic rect may be overwritten");
        } else if (!hasNonWhitePixel) {
          console.warn("[Stage 11] WARNING: All pixels are white! Background image did NOT render on canvas.");
        } else if (hasRedPixel) {
          console.log("[Stage 11] Red diagnostic rect IS visible on canvas. Rendering pipeline working.");
        }
      } catch (e) {
        console.error("[Stage 11] getImageData failed (CORS taint?):", e);
      }

      // Stage 12: Export canvas as PNG
      try {
        canvas.toBlob((blob) => {
          if (blob) {
            const url = URL.createObjectURL(blob);
            console.log("[Stage 12] Canvas exported as blob:", blob.size, "bytes, type:", blob.type);
            console.log("[Stage 12] Blob URL:", url);
            // Create a link to download for comparison
            const a = document.createElement("a");
            a.href = url;
            a.download = "canvas_debug.png";
            // Don't auto-click, just log
            console.log("[Stage 12] Download link ready: canvas_debug.png");
            URL.revokeObjectURL(url);
          } else {
            console.error("[Stage 12] canvas.toBlob returned null!");
          }
        }, "image/png");
      } catch (e) {
        console.error("[Stage 12] canvas.toBlob failed:", e);
      }

      // Also try toDataURL as fallback
      try {
        const dataUrl = canvas.toDataURL("image/png");
        console.log("[Stage 12] canvas.toDataURL() length:", dataUrl.length, "chars");
        // Extract size hint from base64 length (rough: 3 bytes per 4 chars)
        const estimatedBytes = Math.floor(dataUrl.length * 3 / 4);
        console.log("[Stage 12] Estimated PNG size from dataURL:", estimatedBytes, "bytes");
        if (estimatedBytes < 1000) {
          console.warn("[Stage 12] WARNING: Exported PNG is very small (<1KB) - likely blank canvas!");
        }
      } catch (e) {
        console.error("[Stage 12] canvas.toDataURL() failed (CORS taint?):", e);
      }
    });

    hasRenderedRef.current = true;
  }, [backgroundImage, fields, editingFieldId, debug, width, height]);

  const handleClick = useCallback(
    (e: React.MouseEvent<HTMLCanvasElement>) => {
      const canvas = canvasRef.current;
      if (!canvas) return;
      const rect = canvas.getBoundingClientRect();
      const px = e.clientX - rect.left;
      const py = e.clientY - rect.top;
      onClick(px, py);
    },
    [onClick]
  );

  return (
    <canvas
      ref={canvasRef}
      width={width}
      height={height}
      onClick={handleClick}
      style={{ display: "block", cursor: "crosshair" }}
    />
  );
}
