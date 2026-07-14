"use client";

import { useState, useEffect, useCallback, useRef } from "react";
import CanvasViewport from "./CanvasViewport";
import CanvasSurface from "./CanvasSurface";
import HtmlEditorOverlay from "./HtmlEditorOverlay";
import RuntimeDebug from "./RuntimeDebug";
import { hitTest } from "./CanvasHitTester";

export interface LegacyRuntimeField {
  id: string;
  label: string;
  leftPx: number;
  topPx: number;
  widthPx: number;
  heightPx: number;
  type: string;
  required: boolean;
}

export interface LegacyRuntimeDocument {
  backgroundImage: string;
  pageWidth: number;
  pageHeight: number;
  fields: LegacyRuntimeField[];
}

export interface DebugFlags {
  showPageBounds: boolean;
  showFieldRectangles: boolean;
  showFieldIds: boolean;
  showCoordinates: boolean;
  showOriginMarker: boolean;
  showPixelGrid: boolean;
  showHitTestRegions: boolean;
}

export interface ImageLoadInfo {
  status: "idle" | "loading" | "loaded" | "error";
  url: string;
  naturalWidth?: number;
  naturalHeight?: number;
  errorMessage?: string;
}

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:5090";
const RENDER_API_URL = process.env.NEXT_PUBLIC_RENDER_API_URL ?? "http://localhost:5091";

const DEFAULT_DEBUG: DebugFlags = {
  showPageBounds: true,
  showFieldRectangles: true,
  showFieldIds: false,
  showCoordinates: false,
  showOriginMarker: false,
  showPixelGrid: false,
  showHitTestRegions: false,
};

export default function LegacyCanvasRuntime({ templateId = "547" }: { templateId?: string }) {
  const [doc, setDoc] = useState<LegacyRuntimeDocument | null>(null);
  const [bgImage, setBgImage] = useState<HTMLImageElement | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [editingField, setEditingField] = useState<LegacyRuntimeField | null>(
    null
  );
  const [debug, setDebug] = useState<DebugFlags>(DEFAULT_DEBUG);
  const [imgInfo, setImgInfo] = useState<ImageLoadInfo>({ status: "idle", url: "" });
  const [uploading, setUploading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const loadDocument = useCallback((data: LegacyRuntimeDocument, imageBaseUrl?: string) => {
    // Stage 1: Runtime JSON received
    console.log("[Stage 1] Runtime JSON received:", JSON.stringify(data, null, 2));

    if (!data.backgroundImage) {
      console.error("[Stage 2] FATAL: backgroundImage is null or undefined!");
      setError("backgroundImage is missing from API response");
      return;
    }

    console.log("[Stage 2] backgroundImage URL from API:", data.backgroundImage);
    console.log("[Stage 2] pageWidth:", data.pageWidth, "pageHeight:", data.pageHeight);
    console.log("[Stage 2] fieldCount:", data.fields.length);

    setDoc(data);

    // Use imageBaseUrl if provided (upload path → RENDER_API_URL),
    // otherwise default to API_BASE_URL (template path → .NET backend)
    const imageUrl = `${imageBaseUrl ?? API_BASE_URL}${data.backgroundImage}`;
    console.log("[Stage 2] Full image URL:", imageUrl);

    const img = new Image();
    console.log("[Stage 3] new Image() created");

    img.onload = () => {
      console.log("[Stage 6] Image.onload FIRED");
      console.log("[Stage 6] Image properties:", {
        naturalWidth: img.naturalWidth,
        naturalHeight: img.naturalHeight,
        complete: img.complete,
        currentSrc: img.currentSrc,
      });

      if (img.naturalWidth === 0 || img.naturalHeight === 0) {
        console.error("[Stage 6] FATAL: image has zero dimensions!");
      } else {
        console.log("[Stage 6] Image dimensions OK:", img.naturalWidth, "x", img.naturalHeight);
      }

      setImgInfo({
        status: "loaded",
        url: imageUrl,
        naturalWidth: img.naturalWidth,
        naturalHeight: img.naturalHeight,
      });
      setBgImage(img);
    };

    img.onerror = (event: string | Event) => {
      let message = "unknown";
      if (event instanceof Event) {
        message = `Event type: ${event.type}, target src: ${(event.target as HTMLImageElement)?.src || "unknown"}`;
      } else if (typeof event === "string") {
        message = event;
      }
      console.error("[Stage 6] Image.onerror FIRED:", message);
      setImgInfo({
        status: "error",
        url: imageUrl,
        errorMessage: message,
      });
      setBgImage(null);
    };

    img.src = imageUrl;
    console.log("[Stage 4] img.src assigned:", img.src);
    console.log("[Stage 5] Image.complete immediately after src set:", img.complete);
    if (img.complete) {
      console.log("[Stage 5] Image already complete (likely cached) - onload may not fire!");
    }
  }, []);

  const handleUpload = useCallback(async (file: File) => {
    setUploading(true);
    setError(null);
    setDoc(null);
    setBgImage(null);
    setImgInfo({ status: "loading", url: file.name });

    try {
      const formData = new FormData();
      formData.append("file", file);

      const res = await fetch(`${RENDER_API_URL}/upload`, {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const errText = await res.text().catch(() => `HTTP ${res.status}`);
        throw new Error(errText);
      }

      const data: LegacyRuntimeDocument = await res.json();
      console.log("[Upload] RenderResponse received:", data);
      loadDocument(data, RENDER_API_URL);
    } catch (err: any) {
      console.error("[Upload] Failed:", err.message);
      setError(`Upload failed: ${err.message}`);
      setImgInfo({ status: "error", url: file.name, errorMessage: err.message });
    } finally {
      setUploading(false);
    }
  }, [loadDocument]);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = ""; // allow re-selecting the same file
    if (file) {
      handleUpload(file);
    }
  }, [handleUpload]);

  useEffect(() => {
    console.log(`[Stage 0] Fetching template ${templateId} from ${API_BASE_URL}/api/legacyruntime/${templateId}`);
    fetch(`${API_BASE_URL}/api/legacyruntime/${templateId}`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return res.json();
      })
      .then((data: LegacyRuntimeDocument) => {
        loadDocument(data);
      })
      .catch((err: Error) => {
        console.error("[Stage 1] API fetch failed:", err.message);
        console.error("[Stage 1] Full error:", err);
        setError(err.message);
      });
  }, [templateId]);

  const handleCanvasClick = useCallback(
    (pageX: number, pageY: number) => {
      if (!doc) return;
      setEditingField(null);

      const hit = hitTest(doc.fields, pageX, pageY);
      if (hit) {
        setEditingField(hit.field);
      }
    },
    [doc]
  );

  const handleEditorSubmit = useCallback(
    (_value: string) => {
      setEditingField(null);
    },
    []
  );

  const handleEditorCancel = useCallback(() => {
    setEditingField(null);
  }, []);

  return (
    <div style={{ width: "100%", height: "100vh", display: "flex", flexDirection: "column" }}>
      {/* Upload toolbar */}
      <div style={{
        padding: "8px 16px",
        background: "#f5f5f5",
        borderBottom: "1px solid #ddd",
        display: "flex",
        alignItems: "center",
        gap: 12,
        font: "13px sans-serif",
        flexShrink: 0,
      }}>
        <span style={{ fontWeight: 600, color: "#555" }}>Template ID:</span>
        <code style={{ background: "#e8e8e8", padding: "2px 8px", borderRadius: 4 }}>{templateId}</code>
        <span style={{ color: "#aaa" }}>|</span>
        <span style={{ fontWeight: 600, color: "#555" }}>Upload ConMas Workbook:</span>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx"
          onChange={handleFileSelect}
          style={{ display: "none" }}
        />
        <button
          onClick={() => fileInputRef.current?.click()}
          disabled={uploading}
          style={{
            padding: "6px 16px",
            background: uploading ? "#ccc" : "#0070f3",
            color: "#fff",
            border: "none",
            borderRadius: 4,
            cursor: uploading ? "not-allowed" : "pointer",
            font: "13px sans-serif",
          }}
        >
          {uploading ? "Uploading…" : "Choose Excel"}
        </button>
        {uploading && <span style={{ color: "#888" }}>Processing workbook…</span>}
      </div>

      {error && (
        <div style={{ padding: 16, color: "#cc0000", font: "13px sans-serif", background: "#fff0f0", borderBottom: "1px solid #fcc" }}>
          Error: {error}
        </div>
      )}

      {!doc && (
        <div style={{
          flex: 1,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          color: "#666",
          font: "14px sans-serif",
          background: "#fafafa",
        }}>
          {uploading ? (
            <div style={{ textAlign: "center" }}>
              <div style={{ marginBottom: 8 }}>⏳</div>
              <div>Uploading and processing workbook…</div>
            </div>
          ) : (
            <div style={{ textAlign: "center" }}>
              <div style={{ marginBottom: 8, fontSize: 24 }}>📄</div>
              <div>Loading runtime document (Template {templateId})</div>
              <div style={{ marginTop: 12, color: "#999", fontSize: 12 }}>
                or upload a ConMas Excel workbook above
              </div>
            </div>
          )}
        </div>
      )}

      {doc && (<>
        <CanvasViewport width={doc.pageWidth} height={doc.pageHeight}>
          <CanvasSurface
            width={doc.pageWidth}
            height={doc.pageHeight}
            backgroundImage={bgImage}
            fields={doc.fields}
            editingFieldId={editingField?.id ?? null}
            debug={debug}
            onClick={handleCanvasClick}
          />
          {editingField && (
            <HtmlEditorOverlay
              field={editingField}
              onSubmit={handleEditorSubmit}
              onCancel={handleEditorCancel}
            />
          )}
        </CanvasViewport>
        <RuntimeDebug
          flags={debug}
          onChange={setDebug}
          fieldCount={doc.fields.length}
          pageWidth={doc.pageWidth}
          pageHeight={doc.pageHeight}
        />
        <div style={{
          padding: "8px 16px",
          background: imgInfo.status === "loaded" ? "#e6ffe6" : imgInfo.status === "error" ? "#ffe6e6" : imgInfo.status === "loading" ? "#fff8e1" : "#f0f0f0",
          borderTop: "1px solid #ccc",
          font: "12px monospace",
          color: "#333",
        }}>
          <strong>Image Load Status:</strong> {imgInfo.status.toUpperCase()}
          {imgInfo.url && <> | URL: {imgInfo.url}</>}
          {imgInfo.naturalWidth && <> | {imgInfo.naturalWidth}x{imgInfo.naturalHeight}</>}
          {imgInfo.errorMessage && <> | Error: {imgInfo.errorMessage}</>}
        </div>
      </>)}
    </div>
  );
}
