"use client";

import { useState, useEffect, useRef, useCallback, useMemo } from "react";
import { ExcelRenderer, ExcelGrid, OverlayRectangles } from "@/components/ExcelRenderer";
import { fetchTemplate, TEMPLATE_IDS } from "@/services/templateService";
import { generateOverlays, exportOverlays, exportOverlaysBrief } from "@/services/overlayEngine";
import { RuntimeCanvas, RuntimeInspector, useRuntimeState } from "@/components/Runtime";
import type { TemplateModel } from "@/types/template";
import type { OverlayCollection, OverlayModel } from "@/types/overlay";
import type { GridDimensions } from "@/components/ExcelRenderer/types";
import { buildGridDimensions, cellRef, getCellValue, getCellStyle, cumulativeColWidth, cumulativeRowHeight } from "@/components/ExcelRenderer/helpers";
import { DebugOverlay } from "@/components/CompareView/DebugOverlay";
import { CellInspector } from "@/components/CompareView/CellInspector";

type CompareMode = "side-by-side" | "overlay";

interface DebugToggles {
  gridLines: boolean;
  cellCoords: boolean;
  mergeBounds: boolean;
  imageBounds: boolean;
  origin: boolean;
  printArea: boolean;
  margins: boolean;
}

const DEBUG_LABELS: Record<keyof DebugToggles, string> = {
  gridLines: "Grid Lines",
  cellCoords: "Cell Coordinates",
  mergeBounds: "Merge Bounds",
  imageBounds: "Image Bounds",
  origin: "Origin",
  printArea: "Print Area",
  margins: "Margins",
};

const ALL_DEBUG_KEYS = Object.keys(DEBUG_LABELS) as (keyof DebugToggles)[];

export default function ComparePage() {
  const [templateId, setTemplateId] = useState<number>(546);
  const [template, setTemplate] = useState<TemplateModel | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Comparison mode
  const [compareMode, setCompareMode] = useState<CompareMode>("side-by-side");
  const [overlayOpacity, setOverlayOpacity] = useState(50);

  // PDF screenshot
  const [pdfImage, setPdfImage] = useState<string | null>(null);
  const [pdfFile, setPdfFile] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isDragOver, setIsDragOver] = useState(false);

  // Zoom
  const [zoom, setZoom] = useState(100);

  // Debug overlays
  const [showDebug, setShowDebug] = useState(false);
  const [debugToggles, setDebugToggles] = useState<DebugToggles>({
    gridLines: false,
    cellCoords: false,
    mergeBounds: false,
    imageBounds: false,
    origin: false,
    printArea: false,
    margins: false,
  });

  // Cell inspector
  const [selectedCell, setSelectedCell] = useState<{ col: number; row: number } | null>(null);
  const [hoveredCell, setHoveredCell] = useState<{ col: number; row: number } | null>(null);
  const [showInspector, setShowInspector] = useState(false);

  // Overlay engine
  const [showOverlays, setShowOverlays] = useState(false);
  const [selectedOverlay, setSelectedOverlay] = useState<OverlayModel | null>(null);
  const overlays = useMemo(() => template ? generateOverlays(template) : null, [template]);

  // Runtime state
  const runtimeState = useRuntimeState();
  const [showRuntime, setShowRuntime] = useState(false);

  // Metadata
  const [showMeas, setShowMeas] = useState(false);

  // Fetch template
  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    fetchTemplate(templateId)
      .then((data) => {
        if (!cancelled) {
          setTemplate(data);
          setLoading(false);
          setSelectedCell(null);
          setHoveredCell(null);
        }
      })
      .catch((err: Error) => {
        if (!cancelled) {
          setError(err.message);
          setLoading(false);
        }
      });
    return () => { cancelled = true; };
  }, [templateId]);

  const ps = template?.pageSetup;
  const dims: GridDimensions | null = template ? buildGridDimensions(template) : null;
  const totalWidthPt = dims?.totalWidthPt ?? 0;
  const totalHeightPt = dims?.totalHeightPt ?? 0;

  // Cell interaction handlers
  const handleCellClick = useCallback((col: number, row: number) => {
    setSelectedCell({ col, row });
    setShowInspector(true);
  }, []);

  const handleCellHover = useCallback((col: number | null, row: number | null) => {
    if (col !== null && row !== null) {
      setHoveredCell({ col, row });
    } else {
      setHoveredCell(null);
    }
  }, []);

  // PDF upload handlers
  const handleFileSelect = useCallback((file: File) => {
    if (!file.type.startsWith("image/")) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      setPdfImage(e.target?.result as string);
      setPdfFile(file.name);
    };
    reader.readAsDataURL(file);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFileSelect(file);
  }, [handleFileSelect]);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  }, []);

  const handleDragLeave = useCallback(() => {
    setIsDragOver(false);
  }, []);

  const clearPdf = useCallback(() => {
    setPdfImage(null);
    setPdfFile(null);
  }, []);

  // Toggle a debug layer
  const toggleDebug = useCallback((key: keyof DebugToggles) => {
    setDebugToggles((prev) => ({ ...prev, [key]: !prev[key] }));
  }, []);

  // Toggle all debug
  const toggleAllDebug = useCallback(() => {
    const next = !showDebug;
    setShowDebug(next);
    if (next) {
      // Enable all debug overlays
      setDebugToggles({
        gridLines: true,
        cellCoords: true,
        mergeBounds: true,
        imageBounds: true,
        origin: true,
        printArea: true,
        margins: true,
      });
    } else {
      // Disable all
      setDebugToggles({
        gridLines: false,
        cellCoords: false,
        mergeBounds: false,
        imageBounds: false,
        origin: false,
        printArea: false,
        margins: false,
      });
    }
  }, [showDebug]);

  return (
    <div className="min-h-screen bg-gray-100">
      {/* ── Header ── */}
      <header className="bg-white shadow-sm border-b border-gray-200 sticky top-0 z-30">
        <div className="max-w-7xl mx-auto px-4 py-2.5 flex items-center justify-between flex-wrap gap-2">
          <h1 className="text-base font-semibold text-gray-800 whitespace-nowrap">
            Excel Renderer — Visual Validation
          </h1>
          <div className="flex items-center gap-2 flex-wrap">
            {/* Template Selector */}
            <select
              className="border border-gray-300 rounded px-2 py-1 text-xs"
              value={templateId}
              onChange={(e) => setTemplateId(Number(e.target.value))}
            >
              {TEMPLATE_IDS.map((id) => (
                <option key={id} value={id}>Template {id}</option>
              ))}
            </select>

            {/* Mode Toggle */}
            <div className="flex rounded border border-gray-300 overflow-hidden text-xs">
              <button
                className={`px-2.5 py-1 ${compareMode === "side-by-side" ? "bg-blue-500 text-white" : "bg-white text-gray-600 hover:bg-gray-50"}`}
                onClick={() => setCompareMode("side-by-side")}
              >
                Side-by-Side
              </button>
              <button
                className={`px-2.5 py-1 ${compareMode === "overlay" ? "bg-blue-500 text-white" : "bg-white text-gray-600 hover:bg-gray-50"}`}
                onClick={() => setCompareMode("overlay")}
              >
                Overlay
              </button>
            </div>

            {/* Overlay Opacity */}
            {compareMode === "overlay" && (
              <div className="flex items-center gap-1.5 text-xs text-gray-500">
                <span>Opacity</span>
                <input
                  type="range"
                  min={0}
                  max={100}
                  value={overlayOpacity}
                  onChange={(e) => setOverlayOpacity(Number(e.target.value))}
                  className="w-16"
                  title="Overlay opacity"
                />
                <span className="w-7 text-right font-mono">{overlayOpacity}%</span>
              </div>
            )}

            {/* Debug Toggle */}
            <button
              className={`px-2.5 py-1 text-xs rounded border ${showDebug ? "bg-amber-100 border-amber-300 text-amber-700" : "bg-white border-gray-300 text-gray-600 hover:bg-gray-50"}`}
              onClick={toggleAllDebug}
            >
              {showDebug ? "Debug: ON" : "Debug"}
            </button>

            {/* Inspector Toggle */}
            <button
              className={`px-2.5 py-1 text-xs rounded border ${showInspector ? "bg-purple-100 border-purple-300 text-purple-700" : "bg-white border-gray-300 text-gray-600 hover:bg-gray-50"}`}
              onClick={() => setShowInspector((prev) => !prev)}
            >
              {showInspector ? "Inspector: ON" : "Inspector"}
            </button>

            {/* Overlays Toggle */}
            <button
              className={`px-2.5 py-1 text-xs rounded border ${showOverlays ? "bg-emerald-100 border-emerald-300 text-emerald-700" : "bg-white border-gray-300 text-gray-600 hover:bg-gray-50"}`}
              onClick={() => setShowOverlays((prev) => !prev)}
            >
              Overlays
            </button>

            {/* Runtime Toggle */}
            <button
              className={`px-2.5 py-1 text-xs rounded border ${showRuntime ? "bg-indigo-100 border-indigo-300 text-indigo-700" : "bg-white border-gray-300 text-gray-600 hover:bg-gray-50"}`}
              onClick={() => setShowRuntime((prev) => !prev)}
            >
              Runtime
            </button>

            {/* Print PDF Button */}
            <button
              className="px-2.5 py-1 text-xs rounded border bg-blue-500 text-white border-blue-500 hover:bg-blue-600"
              onClick={() => window.print()}
              title="Print or Export PDF — press Ctrl+P or click to open print dialog"
            >
              Print PDF
            </button>

            {/* Zoom */}
            <span className="text-xs text-gray-400">{zoom}%</span>
            <input
              type="range"
              min={25}
              max={200}
              step={5}
              value={zoom}
              onChange={(e) => setZoom(Number(e.target.value))}
              className="w-16"
              title="Zoom"
            />
          </div>
        </div>

        {/* Debug Sub-header */}
        {showDebug && (
          <div className="max-w-7xl mx-auto px-4 pb-2 flex flex-wrap items-center gap-1.5">
            <span className="text-[10px] text-amber-600 font-medium uppercase tracking-wider mr-1">Debug:</span>
            {ALL_DEBUG_KEYS.map((key) => (
              <button
                key={key}
                className={`px-1.5 py-0.5 text-[10px] rounded ${debugToggles[key] ? "bg-amber-200 text-amber-800" : "bg-gray-100 text-gray-500"}`}
                onClick={() => toggleDebug(key)}
              >
                {DEBUG_LABELS[key]}
              </button>
            ))}
          </div>
        )}
      </header>

      <main className="max-w-7xl mx-auto p-4">
        {/* ── SIDE-BY-SIDE MODE ── */}
        {compareMode === "side-by-side" && (
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Left: PDF */}
            <section>
              <h2 className="text-xs font-medium text-gray-500 uppercase tracking-wide mb-2 flex items-center justify-between">
                <span>Microsoft Excel PDF Export</span>
                {pdfFile && (
                  <button onClick={clearPdf} className="text-red-400 hover:text-red-600 text-[10px]">
                    Clear
                  </button>
                )}
              </h2>
              <PDFPanel
                pdfImage={pdfImage}
                pdfFile={pdfFile}
                isDragOver={isDragOver}
                onDrop={handleDrop}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onFileSelect={handleFileSelect}
                fileInputRef={fileInputRef}
                zoom={zoom}
              />
            </section>

            {/* Right: Renderer */}
            <section>
              <h2 className="text-xs font-medium text-gray-500 uppercase tracking-wide mb-2">
                Browser Renderer {loading && <span className="text-blue-400 text-xs">(loading…)</span>}
              </h2>
              <div
                className="bg-white rounded-lg shadow-sm border border-gray-200 overflow-auto relative"
                style={{ maxHeight: "calc(100vh - 180px)" }}
              >
                {error && (
                  <div className="p-4 text-red-500 text-xs">Error: {error}</div>
                )}
                {!error && template && (
                  <div
                    className="flex justify-center p-6"
                    style={{ transform: `scale(${zoom / 100})`, transformOrigin: "top center" }}
                  >
                    {showInspector || showDebug || showOverlays || showRuntime ? (
                      <ExcelGridWrapper
                        template={template}
                        showDebug={showDebug}
                        showOverlays={showOverlays}
                        showRuntime={showRuntime}
                        debugToggles={debugToggles}
                        dims={dims}
                        overlayCollection={overlays}
                        runtimeValues={runtimeState.values}
                        onCellClick={handleCellClick}
                        onCellHover={handleCellHover}
                        onOverlayClick={(o) => setSelectedOverlay(o)}
                        onRuntimeChange={runtimeState.setValue}
                      />
                    ) : (
                      <ExcelRenderer template={template} />
                    )}
                  </div>
                )}
                {!error && !template && !loading && (
                  <div className="p-8 text-center text-gray-400 text-xs">No template loaded</div>
                )}
                {loading && !template && (
                  <div className="p-8 text-center text-gray-300 text-xs">Loading…</div>
                )}
              </div>
            </section>
          </div>
        )}

        {/* ── OVERLAY MODE ── */}
        {compareMode === "overlay" && (
          <div className="flex flex-col items-center gap-4">
            <h2 className="text-xs font-medium text-gray-500 uppercase tracking-wide">
              Overlay Comparison — Browser Renderer over PDF
            </h2>

            {!pdfImage && (
              <div className="w-full max-w-md">
                <PDFPanel
                  pdfImage={pdfImage}
                  pdfFile={pdfFile}
                  isDragOver={isDragOver}
                  onDrop={handleDrop}
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onFileSelect={handleFileSelect}
                  fileInputRef={fileInputRef}
                  zoom={zoom}
                />
              </div>
            )}

            {template && (
              <div
                className="relative"
                style={{
                  width: ps ? `${(ps.paperWidthPt) * (zoom / 100)}pt` : "auto",
                }}
              >
                {/* PDF Image (bottom layer) */}
                {pdfImage && (
                  <div
                    className="rounded-sm overflow-hidden"
                    style={{
                      position: "relative",
                      width: ps ? `${ps.paperWidthPt}pt` : "auto",
                      transform: `scale(${zoom / 100})`,
                      transformOrigin: "top left",
                    }}
                  >
                    <img
                      src={pdfImage}
                      alt="Excel PDF"
                      style={{
                        width: "100%",
                        height: "auto",
                        display: "block",
                      }}
                    />
                  </div>
                )}

                {/* Renderer (top layer, opacity controlled) */}
                {template && (
                  <div
                    className="absolute inset-0 flex items-start justify-center"
                    style={{
                      opacity: overlayOpacity / 100,
                      transform: `scale(${zoom / 100})`,
                      transformOrigin: "top left",
                      ...(pdfImage ? {} : { position: "relative", transform: `scale(${zoom / 100})`, transformOrigin: "top center" }),
                    }}
                  >
                    {showInspector || showDebug || showOverlays || showRuntime ? (
                      <ExcelGridWrapper
                        template={template}
                        showDebug={showDebug}
                        showOverlays={showOverlays}
                        showRuntime={showRuntime}
                        debugToggles={debugToggles}
                        dims={dims}
                        overlayCollection={overlays}
                        runtimeValues={runtimeState.values}
                        onCellClick={handleCellClick}
                        onCellHover={handleCellHover}
                        onOverlayClick={(o) => setSelectedOverlay(o)}
                        onRuntimeChange={runtimeState.setValue}
                      />
                    ) : (
                      <ExcelRenderer template={template} />
                    )}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ── Bottom Panels ── */}
        <div className="mt-6 grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Cell Inspector */}
          {showInspector && template && (
            <section className="bg-white rounded-lg shadow-sm border border-gray-200 p-3">
              <h2 className="text-xs font-medium text-gray-500 uppercase tracking-wide mb-2">
                Cell Inspector
              </h2>
              <CellInspector
                template={template}
                selectedCol={selectedCell?.col ?? null}
                selectedRow={selectedCell?.row ?? null}
              />
            </section>
          )}

          {/* Coordinate Inspector */}
          {hoveredCell && template && (
            <section className="bg-white rounded-lg shadow-sm border border-gray-200 p-3">
              <h2 className="text-xs font-medium text-gray-500 uppercase tracking-wide mb-2">
                Coordinate Inspector
              </h2>
              <CoordinateInspector
                template={template}
                col={hoveredCell.col}
                row={hoveredCell.row}
              />
            </section>
          )}

          {/* Overlay Inspector */}
          {showOverlays && selectedOverlay && (
            <section className="bg-white rounded-lg shadow-sm border border-gray-200 p-3">
              <h2 className="text-xs font-medium text-gray-500 uppercase tracking-wide mb-2">
                Overlay Inspector
              </h2>
              <OverlayInspector overlay={selectedOverlay} />
            </section>
          )}

          {/* Runtime Inspector */}
          {showRuntime && overlays && (
            <section className="bg-white rounded-lg shadow-sm border border-gray-200 p-3">
              <h2 className="text-xs font-medium text-gray-500 uppercase tracking-wide mb-2">
                Runtime Inspector
              </h2>
              <RuntimeInspector
                overlayCollection={overlays}
                values={runtimeState.values}
                dirty={runtimeState.dirty}
                fieldCount={Object.keys(runtimeState.values).length}
                onCopyJson={() => navigator.clipboard.writeText(runtimeState.exportJson())}
                onReset={runtimeState.reset}
              />
            </section>
          )}

          {/* Overlay Export */}
          {showOverlays && overlays && overlays.overlays.length > 0 && (
            <section className="bg-white rounded-lg shadow-sm border border-gray-200 p-3">
              <h2 className="text-xs font-medium text-gray-500 uppercase tracking-wide mb-2">
                Overlay Engine ({overlays.overlays.length} overlays)
              </h2>
              <div className="flex flex-wrap gap-2">
                <button
                  className="px-2 py-1 text-[10px] rounded bg-gray-100 text-gray-600 hover:bg-gray-200 font-mono"
                  onClick={() => {
                    const json = exportOverlays(overlays);
                    navigator.clipboard.writeText(json);
                  }}
                  title="Export as JSON to clipboard"
                >
                  Copy JSON
                </button>
                <button
                  className="px-2 py-1 text-[10px] rounded bg-gray-100 text-gray-600 hover:bg-gray-200 font-mono"
                  onClick={() => {
                    const csv = exportOverlaysBrief(overlays);
                    navigator.clipboard.writeText(csv);
                  }}
                  title="Export as CSV to clipboard"
                >
                  Copy CSV
                </button>
              </div>
              {/* Quick stats */}
              <div className="mt-2 flex flex-wrap gap-2 text-[10px] text-gray-500">
                {countByType(overlays.overlays).map(([type, count]) => (
                  <span key={type} className="bg-gray-50 px-1.5 py-0.5 rounded">
                    {type}: {count}
                  </span>
                ))}
              </div>
            </section>
          )}

          {/* Metadata Panel */}
          {template && (
            <section className={`bg-white rounded-lg shadow-sm border border-gray-200 p-4 ${showInspector && hoveredCell || showOverlays ? "" : "lg:col-span-2"}`}>
              <h2 className="text-xs font-medium text-gray-500 uppercase tracking-wide mb-3">
                Template Metadata
              </h2>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-xs">
                <div>
                  <span className="text-gray-400">Paper</span>
                  <p className="font-mono text-[11px] mt-0.5">{ps?.paperWidthPt} × {ps?.paperHeightPt} pt</p>
                </div>
                <div>
                  <span className="text-gray-400">Orientation</span>
                  <p className="font-mono text-[11px] mt-0.5">{ps?.orientation}</p>
                </div>
                <div>
                  <span className="text-gray-400">Print Area</span>
                  <p className="font-mono text-[11px] mt-0.5">{printAreaSummary(template)}</p>
                </div>
                <div>
                  <span className="text-gray-400">Centering</span>
                  <p className="font-mono text-[11px] mt-0.5">H={ps?.centerHorizontally ? "✓" : "✗"} V={ps?.centerVertically ? "✓" : "✗"}</p>
                </div>
                <div>
                  <span className="text-gray-400">Margins (in)</span>
                  <p className="font-mono text-[11px] mt-0.5">L={ps?.marginLeftIn.toFixed(4)} R={ps?.marginRightIn.toFixed(4)} T={ps?.marginTopIn.toFixed(4)} B={ps?.marginBottomIn.toFixed(4)}</p>
                </div>
                <div>
                  <span className="text-gray-400">Columns</span>
                  <p className="font-mono text-[11px] mt-0.5">{template.columnWidths.length} cols, {template.rowHeights.length} rows</p>
                </div>
                <div>
                  <span className="text-gray-400">Content</span>
                  <p className="font-mono text-[11px] mt-0.5">{Object.keys(template.cellValues).length} values, {template.mergedCells.length} merges, {template.comments.length} comments</p>
                </div>
                <div>
                  <span className="text-gray-400">Total Grid</span>
                  <p className="font-mono text-[11px] mt-0.5">{totalWidthPt.toFixed(1)} × {totalHeightPt.toFixed(1)} pt</p>
                </div>
              </div>

              {/* Measurements toggle */}
              <button
                className="mt-3 text-[10px] text-blue-500 hover:text-blue-700"
                onClick={() => setShowMeas((prev) => !prev)}
              >
                {showMeas ? "Hide" : "Show"} Column/Row Measurements
              </button>

              {showMeas && (
                <div className="mt-3 border-t border-gray-100 pt-3">
                  <h3 className="text-[10px] font-medium text-gray-500 uppercase tracking-wide mb-2">Column Widths</h3>
                  <div className="flex flex-wrap gap-1 text-[10px] font-mono text-gray-600 mb-3">
                    {template.columnWidths.map((w, i) => (
                      <span key={i} className="bg-gray-50 px-1.5 py-0.5 rounded">{String.fromCharCode(65 + i)}={w.toFixed(1)}pt</span>
                    ))}
                  </div>
                  <h3 className="text-[10px] font-medium text-gray-500 uppercase tracking-wide mb-2">Row Heights</h3>
                  <div className="flex flex-wrap gap-1 text-[10px] font-mono text-gray-600">
                    {template.rowHeights.map((h, i) => (
                      <span key={i} className="bg-gray-50 px-1.5 py-0.5 rounded">{i + 1}={h.toFixed(1)}pt</span>
                    ))}
                  </div>
                </div>
              )}
            </section>
          )}
        </div>
      </main>
    </div>
  );
}

// ── Sub-components ──

/** Wraps ExcelGrid with debug overlay, overlay rectangles, and runtime canvas */
function ExcelGridWrapper({
  template,
  showDebug,
  showOverlays,
  showRuntime,
  debugToggles,
  dims,
  overlayCollection,
  runtimeValues,
  onCellClick,
  onCellHover,
  onOverlayClick,
  onRuntimeChange,
}: {
  template: TemplateModel;
  showDebug: boolean;
  showOverlays: boolean;
  showRuntime: boolean;
  debugToggles: DebugToggles;
  dims: GridDimensions | null;
  overlayCollection: OverlayCollection | null;
  runtimeValues: Record<string, string | boolean | null>;
  onCellClick: (col: number, row: number) => void;
  onCellHover: (col: number | null, row: number | null) => void;
  onOverlayClick?: (overlay: OverlayModel) => void;
  onRuntimeChange: (overlayId: string, value: string | boolean | null) => void;
}) {
  const { pageSetup } = template;
  const pageStyle: React.CSSProperties = {
    width: `${pageSetup.paperWidthPt}pt`,
    height: `${pageSetup.paperHeightPt}pt`,
    padding: `${pageSetup.marginTopIn * 72}pt ${pageSetup.marginRightIn * 72}pt ${pageSetup.marginBottomIn * 72}pt ${pageSetup.marginLeftIn * 72}pt`,
    backgroundColor: "#ffffff",
    boxSizing: "border-box",
    overflow: "hidden",
    boxShadow: "0 1px 4px rgba(0,0,0,0.15)",
  };
  const gridWrapperStyle: React.CSSProperties = {
    position: "relative",
    display: "inline-block",
    marginLeft: pageSetup.centerHorizontally ? "auto" : undefined,
    marginRight: pageSetup.centerHorizontally ? "auto" : undefined,
  };

  return (
    <div style={pageStyle}>
      <div style={gridWrapperStyle}>
        <ExcelGrid
          template={template}
          onCellClick={onCellClick}
          onCellHover={onCellHover}
        />
        {showDebug && (
          <DebugOverlay
            template={template}
            showGridLines={debugToggles.gridLines}
            showCellCoords={debugToggles.cellCoords}
            showMergeBounds={debugToggles.mergeBounds}
            showImageBounds={debugToggles.imageBounds}
            showOrigin={debugToggles.origin}
            showPrintArea={debugToggles.printArea}
            showMargins={debugToggles.margins}
          />
        )}
        <OverlayRectangles
          template={template}
          showOverlays={showOverlays}
          onOverlayClick={onOverlayClick}
        />
        {showRuntime && overlayCollection && dims && (
          <RuntimeCanvas
            overlayCollection={overlayCollection}
            runtimeValues={runtimeValues}
            onValueChange={onRuntimeChange}
            widthPt={dims.totalWidthPt}
            heightPt={dims.totalHeightPt}
          />
        )}
      </div>
    </div>
  );
}

/** Drag-drop PDF image panel */
function PDFPanel({
  pdfImage,
  pdfFile,
  isDragOver,
  onDrop,
  onDragOver,
  onDragLeave,
  onFileSelect,
  fileInputRef,
  zoom,
}: {
  pdfImage: string | null;
  pdfFile: string | null;
  isDragOver: boolean;
  onDrop: (e: React.DragEvent) => void;
  onDragOver: (e: React.DragEvent) => void;
  onDragLeave: () => void;
  onFileSelect: (file: File) => void;
  fileInputRef: React.RefObject<HTMLInputElement | null>;
  zoom: number;
}) {
  if (!pdfImage) {
    return (
      <div
        className={`rounded-lg border-2 border-dashed flex items-center justify-center cursor-pointer transition-colors ${
          isDragOver
            ? "border-blue-400 bg-blue-50"
            : "border-gray-300 bg-gray-50 hover:border-gray-400 hover:bg-gray-100"
        }`}
        style={{ minHeight: 300 }}
        onDrop={onDrop}
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
        onClick={() => fileInputRef.current?.click()}
      >
        <div className="text-center p-6">
          <svg className="mx-auto h-10 w-10 mb-2 text-gray-300" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z" />
          </svg>
          <p className="text-sm text-gray-400">Drop Excel PDF screenshot here</p>
          <p className="text-xs text-gray-300 mt-1">or click to browse (PNG, JPEG, WebP)</p>
        </div>
        <input
          ref={fileInputRef}
          type="file"
          accept="image/png,image/jpeg,image/webp"
          className="hidden"
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) onFileSelect(file);
            e.target.value = "";
          }}
        />
      </div>
    );
  }

  return (
    <div className="rounded-lg border border-gray-200 overflow-hidden bg-white">
      <div className="bg-gray-50 px-3 py-1.5 text-[10px] text-gray-500 flex items-center justify-between border-b border-gray-200">
        <span>PDF Screenshot {pdfFile && <span className="font-mono">({pdfFile})</span>}</span>
        <button onClick={(e) => { e.stopPropagation(); }} className="text-red-400 hover:text-red-600 text-[10px]">× Remove</button>
      </div>
      <div
        className="flex justify-center p-4"
        style={{ transform: `scale(${zoom / 100})`, transformOrigin: "top center" }}
      >
        <img
          src={pdfImage}
          alt="Excel PDF Screenshot"
          className="max-w-full h-auto shadow-sm"
          style={{ objectFit: "contain" }}
        />
      </div>
    </div>
  );
}

/** Small coordinate inspector panel with browser vs Excel comparison */
function CoordinateInspector({
  template,
  col,
  row,
}: {
  template: TemplateModel;
  col: number;
  row: number;
}) {
  const { columnWidths, rowHeights, cellStyles, cellValues, hiddenColumns, hiddenRows } = template;
  const ref = cellRef(col, row);
  const style = getCellStyle(col, row, cellStyles);
  const value = getCellValue(col, row, cellValues);
  const colW = columnWidths[col - 1] ?? 0;
  const rowH = rowHeights[row - 1] ?? 15;

  // Use the same cumulative helpers that position cells in the grid
  const absX = cumulativeColWidth(col, columnWidths, hiddenColumns);
  const absY = cumulativeRowHeight(row, rowHeights, hiddenRows);

  // Template values ARE the Excel reference values
  const excelX = cumulativeColWidth(col, columnWidths, hiddenColumns);
  const excelY = cumulativeRowHeight(row, rowHeights, hiddenRows);
  const excelW = colW;
  const excelH = rowH;

  const diffX = absX - excelX;
  const diffY = absY - excelY;
  const diffW = colW - excelW;
  const diffH = rowH - excelH;

  return (
    <div className="text-xs">
      <div className="flex items-center gap-2 mb-2 pb-2 border-b border-gray-200">
        <span className="font-bold font-mono text-blue-700">{ref}</span>
        <span className="text-gray-400">· {value || <span className="italic text-gray-300">(empty)</span>}</span>
      </div>
      <table className="w-full text-[11px]">
        <thead>
          <tr className="text-gray-400 text-[10px]">
            <th className="text-left font-medium">Property</th>
            <th className="text-right font-medium">Browser</th>
            <th className="text-right font-medium">Excel (Expected)</th>
            <th className="text-right font-medium">Diff</th>
          </tr>
        </thead>
        <tbody>
          <CoordRow label="Left" browser={absX} expected={excelX} unit="pt" />
          <CoordRow label="Top" browser={absY} expected={excelY} unit="pt" />
          <CoordRow label="Width" browser={colW} expected={excelW} unit="pt" />
          <CoordRow label="Height" browser={rowH} expected={excelH} unit="pt" />
        </tbody>
      </table>
      <div className="mt-2 pt-2 border-t border-gray-200 flex flex-wrap gap-3 text-[10px]">
        <span className="text-gray-500">Font: <span className="font-mono text-gray-700">{style.fontName ?? "Calibri"} {style.fontSize ?? 11}pt</span></span>
        <span className="text-gray-500">Fill: <span className="font-mono text-gray-700">{style.fillColor ?? "transparent"}</span></span>
        <span className="text-gray-500">Align: <span className="font-mono text-gray-700">{style.horizontalAlignment ?? "left"}/{style.verticalAlignment ?? "middle"}</span></span>
      </div>
    </div>
  );
}

function CoordRow({ label, browser, expected, unit }: { label: string; browser: number; expected: number; unit: string }) {
  const diff = (browser - expected);
  const diffAbs = Math.abs(diff);
  const isMatch = diffAbs < 0.01;
  return (
    <tr className="border-b border-gray-50 last:border-0">
      <td className="text-gray-400 py-0.5">{label}</td>
      <td className="font-mono text-right py-0.5">{browser.toFixed(1)} {unit}</td>
      <td className="font-mono text-right py-0.5 text-gray-500">{expected.toFixed(1)} {unit}</td>
      <td className={`font-mono text-right py-0.5 ${isMatch ? "text-green-600" : diffAbs < 0.5 ? "text-amber-500" : "text-red-500"}`}>
        {isMatch ? "✓ 0" : `${diff > 0 ? "+" : ""}${diff.toFixed(2)}`} {unit}
      </td>
    </tr>
  );
}

/** Overlay inspector panel — shows details of a selected overlay */
function OverlayInspector({ overlay }: { overlay: OverlayModel }) {
  return (
    <div className="text-xs">
      <div className="flex items-center gap-2 mb-2 pb-2 border-b border-gray-200">
        <span className="font-bold font-mono text-sm text-emerald-700">{overlay.id}</span>
        <span className={`text-[10px] px-1.5 py-0.5 rounded font-medium ${getTypeBadgeClass(overlay.type)}`}>
          {overlay.type}
        </span>
      </div>
      <div className="grid grid-cols-2 gap-x-4 gap-y-1.5">
        <div className="text-gray-400">Cell</div>
        <div className="font-mono">{overlay.cell}</div>
        <div className="text-gray-400">Left</div>
        <div className="font-mono">{overlay.leftPt.toFixed(2)} pt</div>
        <div className="text-gray-400">Top</div>
        <div className="font-mono">{overlay.topPt.toFixed(2)} pt</div>
        <div className="text-gray-400">Width</div>
        <div className="font-mono">{overlay.widthPt.toFixed(2)} pt</div>
        <div className="text-gray-400">Height</div>
        <div className="font-mono">{overlay.heightPt.toFixed(2)} pt</div>
        <div className="text-gray-400">Rotation</div>
        <div className="font-mono">{overlay.rotation}°</div>
      </div>
      {/* Metadata */}
      {Object.keys(overlay.metadata).length > 0 && (
        <div className="mt-2 pt-2 border-t border-gray-200">
          <span className="text-gray-400 text-[10px]">Metadata</span>
          <div className="mt-1 font-mono text-[10px] text-gray-600 bg-gray-50 p-1.5 rounded max-h-20 overflow-y-auto">
            {JSON.stringify(overlay.metadata, null, 2)}
          </div>
        </div>
      )}
    </div>
  );
}

function getTypeBadgeClass(type: string): string {
  const colors: Record<string, string> = {
    textbox: "bg-blue-100 text-blue-700",
    signature: "bg-purple-100 text-purple-700",
    checkbox: "bg-green-100 text-green-700",
    date: "bg-amber-100 text-amber-700",
    number: "bg-red-100 text-red-700",
    qr: "bg-pink-100 text-pink-700",
    barcode: "bg-orange-100 text-orange-700",
    image: "bg-violet-100 text-violet-700",
    ocr: "bg-teal-100 text-teal-700",
    unknown: "bg-gray-100 text-gray-700",
  };
  return colors[type] ?? colors.unknown;
}

function countByType(overlays: OverlayModel[]): [string, number][] {
  const counts: Record<string, number> = {};
  for (const o of overlays) {
    counts[o.type] = (counts[o.type] ?? 0) + 1;
  }
  return Object.entries(counts).sort((a, b) => b[1] - a[1]);
}

function printAreaSummary(t: TemplateModel): string {
  const pa = t.printArea;
  const cols = pa.endCol - pa.startCol + 1;
  const rows = pa.endRow - pa.startRow + 1;
  const colLetter = (c: number) => {
    let r = "";
    let n = c;
    while (n > 0) {
      n--;
      r = String.fromCharCode(65 + (n % 26)) + r;
      n = Math.floor(n / 26);
    }
    return r;
  };
  return `${colLetter(pa.startCol)}${pa.startRow}:${colLetter(pa.endCol)}${pa.endRow} (${cols}c×${rows}r)`;
}
