"use client";

import { useState, useRef, useCallback, useEffect, useMemo } from "react";
import type { RuntimeForm, RuntimeField } from "@/types/runtime";
import type { RuntimeState } from "./Runtime/useRuntimeState";
import { RuntimeFormViewer } from "./Runtime/RuntimeFormViewer";

export interface PaperlessDesignerProps {
  runtimeForm: RuntimeForm;
  runtime: RuntimeState;
  templateName: string;
  onReset: () => void;
  onUploadClick: () => void;
}

type ZoomMode = "fit-page" | "fit-width" | "custom";

export function PaperlessDesigner({
  runtimeForm,
  runtime,
  templateName,
  onReset,
  onUploadClick,
}: PaperlessDesignerProps) {
  // ── Multi-page navigation ──
  const [currentPage, setCurrentPage] = useState(0);
  const totalPages = runtimeForm.sheets.length;
  const sheet = runtimeForm.sheets[currentPage] ?? runtimeForm.sheets[0];
  const fields = sheet?.fields ?? [];
  const pageW = sheet?.pageWidthPx ?? runtimeForm.pageWidth;
  const pageH = sheet?.pageHeightPx ?? runtimeForm.pageHeight;

  const goNext = useCallback(() => {
    setCurrentPage((p) => Math.min(p + 1, totalPages - 1));
  }, [totalPages]);

  const goPrev = useCallback(() => {
    setCurrentPage((p) => Math.max(p - 1, 0));
  }, []);

  // Reset page on new form
  useEffect(() => {
    setCurrentPage(0);
  }, [runtimeForm]);

  // ── Selection ──
  const [selectedFieldId, setSelectedFieldId] = useState<string | null>(null);
  const selectedField = useMemo(
    () => fields.find((f) => f.id === selectedFieldId) ?? null,
    [fields, selectedFieldId],
  );

  // ── Field explorer state ──
  const [searchQuery, setSearchQuery] = useState("");
  const [sortBy, setSortBy] = useState<"name" | "cell" | "type">("name");

  const filteredFields = useMemo(() => {
    let result = fields;
    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      result = result.filter(
        (f) =>
          (f.name ?? f.id).toLowerCase().includes(q) ||
          f.cellReference.toLowerCase().includes(q) ||
          f.dataType.toLowerCase().includes(q),
      );
    }
    const sorted = [...result];
    switch (sortBy) {
      case "name":
        sorted.sort((a, b) => (a.name ?? a.id).localeCompare(b.name ?? b.id));
        break;
      case "cell":
        sorted.sort((a, b) => a.cellReference.localeCompare(b.cellReference));
        break;
      case "type":
        sorted.sort((a, b) => a.dataType.localeCompare(b.dataType));
        break;
    }
    return sorted;
  }, [fields, searchQuery, sortBy]);

  // ── Resizable left sidebar ──
  const [leftWidth, setLeftWidth] = useState(280);
  const isResizing = useRef(false);

  const handleResizeStart = useCallback(
    (e: React.MouseEvent) => {
      e.preventDefault();
      isResizing.current = true;
      const startX = e.clientX;
      const startW = leftWidth;
      const onMouseMove = (ev: MouseEvent) => {
        if (!isResizing.current) return;
        setLeftWidth(
          Math.max(200, Math.min(500, startW + (ev.clientX - startX))),
        );
      };
      const onMouseUp = () => {
        isResizing.current = false;
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
      };
      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    },
    [leftWidth],
  );

  // ── Camera model (offsetX, offsetY = paper top-left in screen px, zoom = scale) ──
  const [zoom, setZoom] = useState(1);
  const [offsetX, setOffsetX] = useState(0);
  const [offsetY, setOffsetY] = useState(0);
  const [zoomMode, setZoomMode] = useState<ZoomMode>("fit-page");

  const containerRef = useRef<HTMLDivElement>(null);
  const [cw, setCw] = useState(800);
  const [ch, setCh] = useState(600);

  // Fit Page: compute scale and offset to center paper with padding inside the canvas area
  const fitPage = useCallback(() => {
    if (cw <= 0 || ch <= 0 || pageW <= 0 || pageH <= 0) return;
    const pad = 60;
    const s = Math.max(
      0.1,
      Math.min(8, Math.min((cw - pad * 2) / pageW, (ch - pad * 2) / pageH)),
    );
    setZoom(s);
    setOffsetX((cw - pageW * s) / 2);
    setOffsetY((ch - pageH * s) / 2);
    setZoomMode("fit-page");
  }, [cw, ch, pageW, pageH]);

  // Fit Width
  const fitWidth = useCallback(() => {
    if (cw <= 0 || pageW <= 0) return;
    const pad = 40;
    const s = Math.max(0.1, Math.min(8, (cw - pad * 2) / pageW));
    setZoom(s);
    setOffsetX((cw - pageW * s) / 2);
    setOffsetY((ch - pageH * s) / 3);
    setZoomMode("fit-width");
  }, [cw, ch, pageW, pageH]);

  // Auto-fit on mount and when page/zoomMode changes
  useEffect(() => {
    if (zoomMode === "fit-page") fitPage();
    else if (zoomMode === "fit-width") fitWidth();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cw, ch, zoomMode, currentPage]);

  // ResizeObserver
  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const ro = new ResizeObserver((entries) => {
      for (const e of entries) {
        setCw(e.contentRect.width);
        setCh(e.contentRect.height);
      }
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // ── Zoom toward cursor ──
  const zoomToward = useCallback(
    (newZoom: number, cx: number, cy: number) => {
      const wx = (cx - offsetX) / zoom;
      const wy = (cy - offsetY) / zoom;
      const clampedZoom = Math.max(0.1, Math.min(8, newZoom));
      setZoom(clampedZoom);
      setOffsetX(cx - wx * clampedZoom);
      setOffsetY(cy - wy * clampedZoom);
    },
    [offsetX, offsetY, zoom],
  );

  // ── Zoom to center (for toolbar buttons) ──
  const zoomToCenter = useCallback(
    (newZoom: number) => {
      const cx = cw / 2;
      const cy = ch / 2;
      zoomToward(newZoom, cx, cy);
      setZoomMode("custom");
    },
    [cw, ch, zoomToward],
  );

  // ── Pan state (grab anywhere, no Space key required) ──
  const isPanning = useRef(false);
  const [isGrabbing, setIsGrabbing] = useState(false);
  const panAnchor = useRef({ x: 0, y: 0, ox: 0, oy: 0 });

  const startPan = useCallback(
    (clientX: number, clientY: number) => {
      isPanning.current = true;
      setIsGrabbing(true);
      panAnchor.current = { x: clientX, y: clientY, ox: offsetX, oy: offsetY };
      document.body.style.userSelect = "none";
    },
    [offsetX, offsetY],
  );

  const doPan = useCallback((clientX: number, clientY: number) => {
    if (!isPanning.current) return;
    setOffsetX(panAnchor.current.ox + (clientX - panAnchor.current.x));
    setOffsetY(panAnchor.current.oy + (clientY - panAnchor.current.y));
    setZoomMode("custom");
  }, []);

  const endPan = useCallback(() => {
    isPanning.current = false;
    setIsGrabbing(false);
    document.body.style.userSelect = "";
  }, []);

  const handleMouseDown = useCallback(
    (e: React.MouseEvent) => {
      if (e.button === 0 || e.button === 1) {
        e.preventDefault();
        startPan(e.clientX, e.clientY);
      }
    },
    [startPan],
  );

  const handleMouseMove = useCallback(
    (e: React.MouseEvent) => {
      doPan(e.clientX, e.clientY);
    },
    [doPan],
  );

  const handleMouseUp = useCallback(() => endPan(), [endPan]);

  // ── Wheel zoom (Ctrl/Meta = zoom toward cursor, else = pan) ──
  const handleWheel = useCallback(
    (e: React.WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) {
        setOffsetX((o) => o - e.deltaX);
        setOffsetY((o) => o - e.deltaY);
        setZoomMode("custom");
        return;
      }
      e.preventDefault();
      const rect = containerRef.current?.getBoundingClientRect();
      if (!rect) return;
      const cx = e.clientX - rect.left;
      const cy = e.clientY - rect.top;
      const factor = Math.pow(1.001, -e.deltaY);
      const newZoom = zoom * factor;
      zoomToward(newZoom, cx, cy);
      setZoomMode("custom");
    },
    [zoom, zoomToward],
  );

  // ── Zoom helpers (fine 5% increments instead of preset jumps) ──
  const zoomIn = useCallback(() => {
    zoomToCenter(Math.round((zoom + 0.05) * 100) / 100);
  }, [zoom, zoomToCenter]);

  const zoomOut = useCallback(() => {
    zoomToCenter(Math.round((zoom - 0.05) * 100) / 100);
  }, [zoom, zoomToCenter]);

  const zoom100 = useCallback(() => {
    zoomToCenter(1);
  }, [zoomToCenter]);

  const handleDoubleClick = useCallback(() => fitPage(), [fitPage]);

  // ── Overlay / background toggles ──
  const [showOverlay, setShowOverlay] = useState(true);
  const [showBackground, setShowBackground] = useState(true);

  const cursorClass = isGrabbing ? "grabbing" : "grab";
  const currentZoomPercent = Math.round(zoom * 100);

  return (
    <div className="flex flex-col h-full" style={{ outline: "none" }}>
      {/* ── Toolbar ── */}
      <Toolbar
        templateName={templateName}
        zoomPercent={currentZoomPercent}
        zoomMode={zoomMode}
        currentPage={currentPage}
        totalPages={totalPages}
        onGoPrev={goPrev}
        onGoNext={goNext}
        onReset={onReset}
        onUploadClick={onUploadClick}
        onZoomIn={zoomIn}
        onZoomOut={zoomOut}
        onZoom100={zoom100}
        onFitPage={fitPage}
        onFitWidth={fitWidth}
        showOverlay={showOverlay}
        onToggleOverlay={() => setShowOverlay((v) => !v)}
        showBackground={showBackground}
        onToggleBackground={() => setShowBackground((v) => !v)}
        onPageSelect={setCurrentPage}
      />

      {/* ── Main area ── */}
      <div className="flex flex-1 overflow-hidden ">
        {/* Left sidebar */}
        <div
          style={{ width: leftWidth, minWidth: 200 }}
          className="bg-white border-r border-slate-200 flex flex-col overflow-hidden "
        >
          <FieldExplorer
            fields={filteredFields}
            totalFields={fields.length}
            selectedFieldId={selectedFieldId}
            searchQuery={searchQuery}
            onSearchChange={setSearchQuery}
            sortBy={sortBy}
            onSortChange={setSortBy}
            onFieldSelect={setSelectedFieldId}
          />
        </div>

        <div
          onMouseDown={handleResizeStart}
          className="w-1 cursor-col-resize bg-transparent hover:bg-emerald-400 active:bg-emerald-500 shrink-0 transition-colors"
        />

        {/* ── Infinite workspace ── */}
        <div
          ref={containerRef}
          className="flex-1 overflow-hidden relative select-none "
          style={{
            backgroundColor: "#e6e6e6",
            cursor: cursorClass,
          }}
          onMouseDown={handleMouseDown}
          onMouseMove={handleMouseMove}
          onMouseUp={handleMouseUp}
          onMouseLeave={handleMouseUp}
          onWheel={handleWheel}
          onDoubleClick={handleDoubleClick}
        >
          {/* Grid dots pattern for infinite canvas feel */}
          <div
            style={{
              position: "absolute",
              inset: 0,
              backgroundImage:
                "radial-gradient(circle, rgba(0,0,0,0.08) 1px, transparent 1px)",
              backgroundSize: "24px 24px",
              pointerEvents: "none",
            }}
          />

          {/* Camera transform: unified translate + scale for smooth pan/zoom */}
          <div
            style={{
              position: "absolute",
              transform: `translate(${offsetX}px, ${offsetY}px) scale(${zoom})`,
              transformOrigin: "0 0",
              lineHeight: 0,
              willChange: "transform",
              filter: "drop-shadow(0 4px 24px rgba(0,0,0,0.15))",
            }}
          >
            {/* Paper wrapper with rounded corners + clip */}
            <div
              style={{
                borderRadius: 2,
                overflow: "hidden",
                backgroundColor: "#ffffff",
              }}
            >
              <RuntimeFormViewer
                runtimeForm={runtimeForm}
                runtime={runtime}
                selectedFieldId={selectedFieldId}
                currentPage={currentPage}
                showOverlay={showOverlay}
                showBackground={showBackground}
              />
            </div>
          </div>

          {/* ── Bottom-right overlay: Page navigation + zoom ── */}
          <div
            className="absolute bottom-4 right-4 flex flex-col gap-2 z-40 select-none"
            onClick={(e) => e.stopPropagation()}
          >
            {/* Page navigation */}
            {totalPages > 1 && (
              <div className="bg-white/95 backdrop-blur-sm border border-slate-200 rounded-xl shadow-sm flex items-center gap-1 px-1 py-1">
                <button
                  onClick={goPrev}
                  disabled={currentPage === 0}
                  className="w-8 h-8 flex items-center justify-center rounded-lg hover:bg-slate-100 text-slate-600 disabled:text-slate-300 disabled:hover:bg-transparent transition-colors"
                  title="Previous page"
                >
                  <svg
                    className="w-4 h-4"
                    fill="none"
                    viewBox="0 0 24 24"
                    stroke="currentColor"
                    strokeWidth={2}
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M15 19l-7-7 7-7"
                    />
                  </svg>
                </button>
                <span className="text-xs font-medium text-slate-700 min-w-[60px] text-center select-none">
                  {currentPage + 1} / {totalPages}
                </span>
                <button
                  onClick={goNext}
                  disabled={currentPage >= totalPages - 1}
                  className="w-8 h-8 flex items-center justify-center rounded-lg hover:bg-slate-100 text-slate-600 disabled:text-slate-300 disabled:hover:bg-transparent transition-colors"
                  title="Next page"
                >
                  <svg
                    className="w-4 h-4"
                    fill="none"
                    viewBox="0 0 24 24"
                    stroke="currentColor"
                    strokeWidth={2}
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M9 5l7 7-7 7"
                    />
                  </svg>
                </button>
              </div>
            )}

            {/* Zoom controls */}
            <div className="bg-white/95 backdrop-blur-sm border border-slate-200 rounded-xl shadow-sm flex items-center gap-1 px-1 py-1">
              <button
                onClick={zoomOut}
                className="w-8 h-8 flex items-center justify-center rounded-lg hover:bg-slate-100 text-slate-600 transition-colors"
                title="Zoom out"
              >
                −
              </button>
              <span
                className="min-w-[52px] text-center text-xs font-medium text-slate-700 cursor-pointer hover:bg-slate-100 rounded-lg py-1 transition-colors"
                onClick={fitPage}
                title="Fit Page"
              >
                {currentZoomPercent}%
              </span>
              <button
                onClick={zoomIn}
                className="w-8 h-8 flex items-center justify-center rounded-lg hover:bg-slate-100 text-slate-600 transition-colors"
                title="Zoom in"
              >
                +
              </button>
            </div>
          </div>
        </div>

        {/* Right sidebar */}
        <div
          style={{ width: 280 }}
          className="bg-white border-l border-slate-200 flex flex-col overflow-hidden"
        >
          <FieldPropertiesPanel field={selectedField} />
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════
   Toolbar
   ═══════════════════════════════════════════════ */
interface ToolbarProps {
  templateName: string;
  zoomPercent: number;
  zoomMode: ZoomMode;
  currentPage: number;
  totalPages: number;
  onGoPrev: () => void;
  onGoNext: () => void;
  onReset: () => void;
  onUploadClick: () => void;
  onZoomIn: () => void;
  onZoomOut: () => void;
  onZoom100: () => void;
  onFitPage: () => void;
  onFitWidth: () => void;
  showOverlay: boolean;
  onToggleOverlay: () => void;
  showBackground: boolean;
  onToggleBackground: () => void;
  onPageSelect: (page: number) => void;
}

function Toolbar({
  templateName,
  zoomPercent,
  zoomMode,
  currentPage,
  totalPages,
  onGoPrev,
  onGoNext,
  onReset,
  onUploadClick,
  onZoomIn,
  onZoomOut,
  onZoom100,
  onFitPage,
  onFitWidth,
  showOverlay,
  onToggleOverlay,
  showBackground,
  onToggleBackground,
  onPageSelect,
}: ToolbarProps) {
  return (
    <div className="h-12 bg-white border-b border-slate-200 flex items-center px-3 gap-1 shrink-0 select-none">
      {/* Logo */}
      <div className="flex items-center gap-2 mr-2">
        <div className="w-6 h-6 rounded-md bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center shadow-sm">
          <svg
            className="w-3 h-3 text-white"
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
            strokeWidth={2.5}
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
            />
          </svg>
        </div>
        <span className="text-sm font-semibold text-slate-800">PaperLess</span>
      </div>

      <div className="w-px h-6 bg-slate-200 mx-1" />

      {/* File */}
      <ToolbarButton
        onClick={onReset}
        title="Open Template"
        icon={
          <svg
            className="w-4 h-4"
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
            strokeWidth={2}
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"
            />
          </svg>
        }
      />
      <ToolbarButton
        onClick={onUploadClick}
        title="Upload"
        icon={
          <svg
            className="w-4 h-4"
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
            strokeWidth={2}
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"
            />
          </svg>
        }
      />

      <div className="w-px h-6 bg-slate-200 mx-1" />

      {/* View controls */}
      <ToolbarButton
        onClick={onFitPage}
        title="Fit Page"
        active={zoomMode === "fit-page"}
        icon={
          <svg
            className="w-4 h-4"
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
            strokeWidth={2}
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M4 8V4m0 0h4M4 4l5 5m11-1V4m0 0h-4m4 0l-5 5M4 16v4m0 0h4m-4 0l5-5m11 5v-4m0 4h-4m4 0l-5-5"
            />
          </svg>
        }
      />
      <ToolbarButton
        onClick={onFitWidth}
        title="Fit Width"
        active={zoomMode === "fit-width"}
        icon={
          <svg
            className="w-4 h-4"
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
            strokeWidth={2}
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M5 12h14M5 12l4-4m-4 4l4 4m11-4l-4-4m4 4l-4 4"
            />
          </svg>
        }
      />
      <ToolbarButton
        onClick={onZoomOut}
        title="Zoom Out"
        icon={
          <svg
            className="w-4 h-4"
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
            strokeWidth={2}
          >
            <path strokeLinecap="round" strokeLinejoin="round" d="M20 12H4" />
          </svg>
        }
      />
      <ToolbarButton
        onClick={onZoom100}
        title="Zoom to 100%"
        icon={<span className="text-xs font-semibold">1:1</span>}
      />
      <ToolbarButton
        onClick={onZoomIn}
        title="Zoom In"
        icon={
          <svg
            className="w-4 h-4"
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
            strokeWidth={2}
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M12 4v16m8-8H4"
            />
          </svg>
        }
      />
      <span className="text-xs font-medium text-slate-600 w-12 text-center select-none">
        {zoomPercent}%
      </span>

      <div className="w-px h-6 bg-slate-200 mx-1" />

      {/* Visibility toggles */}
      <button
        onClick={onToggleOverlay}
        title="Show or hide interactive form fields"
        className={`flex items-center gap-1.5 h-8 px-2.5 rounded-md text-xs font-medium transition-colors ${
          showOverlay
            ? "bg-emerald-100 text-emerald-700"
            : "text-slate-400 hover:text-slate-600 hover:bg-slate-100"
        }`}
      >
        <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={showOverlay ? 2 : 1.5}>
          {showOverlay ? (
            <path strokeLinecap="round" strokeLinejoin="round" d="M2.036 12.322a1.012 1.012 0 010-.639C3.423 7.51 7.36 4.5 12 4.5c4.638 0 8.573 3.007 9.963 7.178.07.207.07.431 0 .639C20.577 16.49 16.64 19.5 12 19.5c-4.638 0-8.573-3.007-9.963-7.178z" />
          ) : (
            <>
              <path strokeLinecap="round" strokeLinejoin="round" d="M2.036 12.322a1.012 1.012 0 010-.639C3.423 7.51 7.36 4.5 12 4.5c4.638 0 8.573 3.007 9.963 7.178.07.207.07.431 0 .639C20.577 16.49 16.64 19.5 12 19.5c-4.638 0-8.573-3.007-9.963-7.178z" />
              <path strokeLinecap="round" strokeLinejoin="round" d="M3 3l18 18" />
            </>
          )}
        </svg>
        <span>Fields</span>
      </button>
      <button
        onClick={onToggleBackground}
        title="Show or hide the rendered Excel page"
        className={`flex items-center gap-1.5 h-8 px-2.5 rounded-md text-xs font-medium transition-colors ${
          showBackground
            ? "bg-emerald-100 text-emerald-700"
            : "text-slate-400 hover:text-slate-600 hover:bg-slate-100"
        }`}
      >
        <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={showBackground ? 2 : 1.5}>
          {showBackground ? (
            <path strokeLinecap="round" strokeLinejoin="round" d="M2.036 12.322a1.012 1.012 0 010-.639C3.423 7.51 7.36 4.5 12 4.5c4.638 0 8.573 3.007 9.963 7.178.07.207.07.431 0 .639C20.577 16.49 16.64 19.5 12 19.5c-4.638 0-8.573-3.007-9.963-7.178z" />
          ) : (
            <>
              <path strokeLinecap="round" strokeLinejoin="round" d="M2.036 12.322a1.012 1.012 0 010-.639C3.423 7.51 7.36 4.5 12 4.5c4.638 0 8.573 3.007 9.963 7.178.07.207.07.431 0 .639C20.577 16.49 16.64 19.5 12 19.5c-4.638 0-8.573-3.007-9.963-7.178z" />
              <path strokeLinecap="round" strokeLinejoin="round" d="M3 3l18 18" />
            </>
          )}
        </svg>
        <span>Background</span>
      </button>

      {/* Spacer */}
      <div className="flex-1" />

      {/* Page navigation (desktop toolbar) */}
      {totalPages > 1 && (
        <>
          <button
            onClick={onGoPrev}
            disabled={currentPage === 0}
            className="flex items-center gap-1 px-2 py-1 rounded-md text-xs text-slate-600 hover:bg-slate-100 disabled:text-slate-300 disabled:hover:bg-transparent transition-colors"
          >
            <svg
              className="w-3.5 h-3.5"
              fill="none"
              viewBox="0 0 24 24"
              stroke="currentColor"
              strokeWidth={2}
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M15 19l-7-7 7-7"
              />
            </svg>
            Previous
          </button>

          <span className="text-xs font-medium text-slate-500 min-w-[80px] text-center select-none">
            Page {currentPage + 1} <span className="text-slate-300">of</span>{" "}
            {totalPages}
          </span>

          <button
            onClick={onGoNext}
            disabled={currentPage >= totalPages - 1}
            className="flex items-center gap-1 px-2 py-1 rounded-md text-xs text-slate-600 hover:bg-slate-100 disabled:text-slate-300 disabled:hover:bg-transparent transition-colors"
          >
            Next
            <svg
              className="w-3.5 h-3.5"
              fill="none"
              viewBox="0 0 24 24"
              stroke="currentColor"
              strokeWidth={2}
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M9 5l7 7-7 7"
              />
            </svg>
          </button>

          <div className="w-px h-6 bg-slate-200 mx-1" />
        </>
      )}

      {/* Sheet names as page thumbnails */}
      <PageThumbnails
        currentPage={currentPage}
        totalPages={totalPages}
        onSelect={onPageSelect}
      />

      {/* Template name */}
      <span className="text-xs text-slate-400 truncate max-w-[160px] ml-2">
        {templateName}
      </span>
    </div>
  );
}

/* ── Page Thumbnails in Toolbar ── */
function PageThumbnails({
  currentPage,
  totalPages,
  onSelect,
}: {
  currentPage: number;
  totalPages: number;
  onSelect: (page: number) => void;
}) {
  // Show page numbers as clickable chips
  return (
    <div className="flex items-center gap-0.5 mr-1">
      {Array.from({ length: totalPages }, (_, i) => (
        <button
          key={i}
          onClick={() => onSelect(i)}
          className={`min-w-[24px] h-6 flex items-center justify-center rounded text-[11px] font-medium transition-all duration-150 ${
            i === currentPage
              ? "bg-emerald-100 text-emerald-700 shadow-sm"
              : "text-slate-500 hover:bg-slate-100"
          }`}
        >
          {i + 1}
        </button>
      ))}
    </div>
  );
}

function ToolbarButton({
  onClick,
  title,
  icon,
  active,
}: {
  onClick: () => void;
  title: string;
  icon: React.ReactNode;
  active?: boolean;
}) {
  return (
    <button
      onClick={onClick}
      title={title}
      className={`w-8 h-8 flex items-center justify-center rounded-md transition-colors ${
        active
          ? "bg-emerald-100 text-emerald-700"
          : "text-slate-500 hover:bg-slate-100 hover:text-slate-700"
      }`}
    >
      {icon}
    </button>
  );
}

/* ═══════════════════════════════════════════════
   Field Explorer (Left Sidebar)
   ═══════════════════════════════════════════════ */
interface FieldExplorerProps {
  fields: RuntimeField[];
  totalFields: number;
  selectedFieldId: string | null;
  searchQuery: string;
  onSearchChange: (q: string) => void;
  sortBy: "name" | "cell" | "type";
  onSortChange: (s: "name" | "cell" | "type") => void;
  onFieldSelect: (id: string | null) => void;
}

function FieldExplorer({
  fields,
  totalFields,
  selectedFieldId,
  searchQuery,
  onSearchChange,
  sortBy,
  onSortChange,
  onFieldSelect,
}: FieldExplorerProps) {
  return (
    <>
      <div className="px-3 py-2 border-b border-slate-100">
        <div className="flex items-center justify-between mb-1">
          <span className="text-xs font-semibold text-slate-700 uppercase tracking-wider">
            Fields
          </span>
          <span className="text-[10px] text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded-full">
            {totalFields}
          </span>
        </div>
        <input
          type="text"
          value={searchQuery}
          onChange={(e) => onSearchChange(e.target.value)}
          placeholder="Search fields..."
          className="w-full text-xs px-2 py-1 border border-slate-200 rounded-md focus:outline-none focus:border-emerald-400 bg-slate-50"
        />
        <div className="flex gap-1 mt-1">
          {(["name", "cell", "type"] as const).map((s) => (
            <button
              key={s}
              onClick={() => onSortChange(s)}
              className={`text-[10px] px-1.5 py-0.5 rounded ${
                sortBy === s
                  ? "bg-emerald-100 text-emerald-700 font-medium"
                  : "text-slate-400 hover:text-slate-600"
              }`}
            >
              {s.charAt(0).toUpperCase() + s.slice(1)}
            </button>
          ))}
        </div>
      </div>

      <div className="flex-1 overflow-y-auto">
        {fields.map((field) => (
          <div
            key={field.id}
            onClick={() => onFieldSelect(field.id)}
            className={`px-3 py-2 cursor-pointer border-b border-slate-50 hover:bg-slate-50 transition-colors ${
              field.id === selectedFieldId
                ? "bg-emerald-50 border-l-2 border-l-emerald-500"
                : ""
            }`}
          >
            <div className="flex items-center gap-2">
              <FieldTypeIcon type={field.dataType} />
              <span className="text-xs font-medium text-slate-800 truncate flex-1">
                {field.name ?? field.id}
              </span>
            </div>
            {field.name && (
              <div className="ml-5 mt-0.5">
                <span className="text-xs text-slate-400 font-mono">
                  {field.id}
                </span>
              </div>
            )}
            <div className="flex items-center gap-2 mt-0.5 ml-5">
              <span className="text-[10px] text-slate-400 font-mono">
                {field.cellReference}
              </span>
              <span className="text-[10px] text-slate-400">
                {field.dataType}
              </span>
            </div>
          </div>
        ))}
        {fields.length === 0 && (
          <div className="px-3 py-6 text-center text-xs text-slate-400">
            {totalFields === 0
              ? "No fields detected"
              : "No fields match search"}
          </div>
        )}
      </div>
    </>
  );
}

function FieldTypeIcon({ type }: { type: string }) {
  const iconMap: Record<string, { char: string; color: string }> = {
    text: { char: "T", color: "text-blue-500" },
    number: { char: "#", color: "text-orange-500" },
    date: { char: "D", color: "text-purple-500" },
    checkbox: { char: "\u2713", color: "text-emerald-500" },
    signature: { char: "S", color: "text-rose-500" },
    dropdown: { char: "\u25BC", color: "text-cyan-500" },
    calculated: { char: "\u2211", color: "text-slate-500" },
  };
  const meta = iconMap[type] ?? { char: "?", color: "text-slate-400" };
  return (
    <span
      className={`w-4 h-4 flex items-center justify-center text-[10px] font-bold rounded ${meta.color}`}
      style={{ fontSize: 9, lineHeight: 1 }}
    >
      {meta.char}
    </span>
  );
}

/* ═══════════════════════════════════════════════
   Field Properties Panel (Right Sidebar)
   ═══════════════════════════════════════════════ */
function FieldPropertiesPanel({ field }: { field: RuntimeField | null }) {
  if (!field) {
    return (
      <div className="p-4 text-xs text-slate-400 text-center mt-8">
        Select a field to view properties
      </div>
    );
  }

  const rows: [string, string][] = [
    ["Name", field.name ?? field.id],
    ["Cell", field.cellReference],
    ["Type", field.dataType],
    ["Left (ratio)", field.leftRatio.toFixed(4)],
    ["Top (ratio)", field.topRatio.toFixed(4)],
    ["Right (ratio)", (field.leftRatio + field.widthRatio).toFixed(4)],
    ["Bottom (ratio)", (field.topRatio + field.heightRatio).toFixed(4)],
    ["Width (px)", String(Math.round(field.widthPx))],
    ["Height (px)", String(Math.round(field.heightPx))],
    ["Merge Range", field.mergeRange ?? "\u2014"],
    ["Read Only", field.readOnly ? "Yes" : "No"],
    ["Required", field.required ? "Yes" : "No"],
  ];

  return (
    <div className="flex flex-col h-full">
      <div className="px-3 py-2 border-b border-slate-100">
        <span className="text-xs font-semibold text-slate-700 uppercase tracking-wider">
          Properties
        </span>
      </div>
      <div className="flex-1 overflow-y-auto p-3 space-y-2">
        {rows.map(([label, value]) => (
          <div key={label}>
            <div className="text-[10px] text-slate-400 uppercase tracking-wide">
              {label}
            </div>
            <div className="text-xs text-slate-800 font-mono truncate">
              {value}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
