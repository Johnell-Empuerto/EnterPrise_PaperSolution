"use client";

import { useState, useEffect, useRef } from "react";
import type { RuntimeSheet, RuntimeField } from "@/types/runtime";
import { OverlayRenderer } from "./OverlayRenderer";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:5090";

interface FormPageProps {
  /** The runtime sheet data for this page. */
  sheet: RuntimeSheet;
  /**
   * The background preview URL from the backend.
   * If not provided, falls back to constructing from templateId.
   */
  previewUrl?: string | null;
  /** Template ID for fallback background image URL construction. */
  templateId: string;
  /** Zoom level (0.25–3). */
  zoom: number;
  /** Called when a field is focused. */
  onFieldFocus?: (field: RuntimeField) => void;
  /** Visual debug mode: renders green border around each overlay field. */
  debug?: boolean;
}

/**
 * Renders a single page of the runtime form: background image + overlay fields.
 * The background PNG is rendered as an <img>, and each editable field is
 * positioned on top using absolute positioning with coordinates from the backend.
 *
 * Background image resolution (priority order):
 *   1. sheet.backgroundImage (from runtime) — the exact image used during coordinate computation
 *   2. previewUrl prop (from upload response) — the preview URL returned during upload
 *   3. /preview/page_{templateId}.png — constructed from template ID (backward compat)
 *
 * Dimension validation: On image load, verifies the actual image dimensions match
 * the expected pageWidthPx/pageHeightPx. If they differ, shows a warning overlay.
 */
export function FormPage({ sheet, previewUrl, templateId, zoom, onFieldFocus, debug }: FormPageProps) {
  const [bgError, setBgError] = useState(false);
  const [dimMismatch, setDimMismatch] = useState<string | null>(null);
  const imgRef = useRef<HTMLImageElement>(null);

  // Resolve the background image URL (priority: runtime > upload > fallback)
  const bgUrl = sheet.backgroundImage
    ? `${API_BASE_URL}${sheet.backgroundImage.startsWith("/") ? "" : "/"}${sheet.backgroundImage}`
    : previewUrl
    ? `${API_BASE_URL}${previewUrl.startsWith("/") ? "" : "/"}${previewUrl}`
    : `${API_BASE_URL}/preview/page_${templateId}.png`;

  // ── DEBUG: Log background image info ────────────────────────────────
  console.log("[DEBUG] FormPage render", {
    sheetName: sheet.name,
    pageWidthPx: sheet.pageWidthPx,
    pageHeightPx: sheet.pageHeightPx,
    fieldCount: sheet.fields.length,
    backgroundImage: sheet.backgroundImage,
    previewUrl_prop: previewUrl,
    resolvedBgUrl: bgUrl,
    templateId,
    zoom,
  });

  // Reset error state when template changes
  useEffect(() => {
    setBgError(false);
    setDimMismatch(null);
  }, [previewUrl, templateId]);

  // Validate image dimensions on load
  const handleImageLoad = () => {
    const img = imgRef.current;
    if (!img) return;

    const naturalW = img.naturalWidth;
    const naturalH = img.naturalHeight;
    const expectedW = sheet.pageWidthPx;
    const expectedH = sheet.pageHeightPx;

    console.log("[DEBUG] Image loaded:", {
      url: bgUrl,
      naturalW,
      naturalH,
      expectedW,
      expectedH,
      match: naturalW === expectedW && naturalH === expectedH,
    });

    if (naturalW !== expectedW || naturalH !== expectedH) {
      setDimMismatch(
        `Image dimensions mismatch: expected ${expectedW}x${expectedH}, actual ${naturalW}x${naturalH}. ` +
        `The background image has been stretched to fit, causing overlay misalignment.`
      );
      console.error("[DEBUG] IMAGE DIMENSION MISMATCH — overlay will be misaligned!");
    } else {
      setDimMismatch(null);
      console.log("[DEBUG] Image dimensions MATCH — no stretching.");
    }
  };

  const handleBgError = () => {
    console.error("[DEBUG] Background image failed to load:", { bgUrl });
    setBgError(true);
  };

  return (
    <div
      style={{
        position: "relative",
        width: sheet.pageWidthPx,
        height: sheet.pageHeightPx,
        transform: `scale(${zoom})`,
        transformOrigin: "top left",
      }}
    >
      {/* Debug overlay for dimension mismatch */}
      {dimMismatch && (
        <div
          style={{
            position: "absolute",
            top: 0,
            left: 0,
            right: 0,
            zIndex: 50,
            backgroundColor: "rgba(255, 0, 0, 0.85)",
            color: "white",
            padding: "8px 12px",
            fontSize: "13px",
            fontFamily: "monospace",
            pointerEvents: "none",
          }}
        >
          ⚠ {dimMismatch}
        </div>
      )}

      {/* Background PNG — fills the container exactly */}
      {!bgError ? (
        // eslint-disable-next-line @next/next/no-img-element
        <img
          ref={imgRef}
          src={bgUrl}
          alt={`Sheet: ${sheet.name}`}
          style={{
            display: "block",
            width: sheet.pageWidthPx,
            height: sheet.pageHeightPx,
          }}
          onError={handleBgError}
          onLoad={handleImageLoad}
          draggable={false}
        />
      ) : (
        <div
          className="bg-gray-100 flex items-center justify-center text-gray-400 text-sm"
          style={{ width: sheet.pageWidthPx, height: sheet.pageHeightPx }}
        >
          Background image not available
        </div>
      )}

      {/* Field overlays — positioned absolutely on top of the background */}
      <OverlayRenderer
        fields={sheet.fields}
        onFieldFocus={onFieldFocus}
        scale={zoom}
        pageWidthPx={sheet.pageWidthPx}
        pageHeightPx={sheet.pageHeightPx}
        debug={debug}
      />
    </div>
  );
}
