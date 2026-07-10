"use client";

import { useState, useEffect } from "react";
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
 * The background URL comes from the backend (single source of truth).
 * If previewUrl is provided, it is used directly. Otherwise, falls back
 * to constructing /preview/page_{templateId}.png for backward compatibility.
 */
export function FormPage({ sheet, previewUrl, templateId, zoom, onFieldFocus, debug }: FormPageProps) {
  const [bgError, setBgError] = useState(false);

  // Use the backend-provided previewUrl if available; otherwise fall back
  const bgUrl = previewUrl
    ? `${API_BASE_URL}${previewUrl.startsWith("/") ? "" : "/"}${previewUrl}`
    : `${API_BASE_URL}/preview/page_${templateId}.png`;

  // Reset error state when template changes
  useEffect(() => {
    setBgError(false);
  }, [previewUrl, templateId]);

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
      {/* Background PNG — fills the container exactly */}
      {!bgError ? (
        // eslint-disable-next-line @next/next/no-img-element
        <img
          src={bgUrl}
          alt={`Sheet: ${sheet.name}`}
          style={{
            display: "block",
            width: sheet.pageWidthPx,
            height: sheet.pageHeightPx,
          }}
          onError={() => setBgError(true)}
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
        debug={debug}
      />
    </div>
  );
}
