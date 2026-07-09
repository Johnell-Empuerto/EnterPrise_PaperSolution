"use client";

import { useState, useRef, useEffect, useCallback, type FormEvent } from "react";
import Image from "next/image";
import { DesignerController } from "@/lib/designerController";
import type {
  FormDefinition,
  SheetDefinition,
  ClusterDefinition,
} from "@/lib/formDefinition";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:5090";

const controller = new DesignerController();

interface ApiResponse<T> {
  success: boolean;
  message: string;
  data?: T;
}

export default function Home() {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [message, setMessage] = useState<string | null>(null);
  const [isError, setIsError] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [zoom, setZoom] = useState(0.5);
  const [dirty, setDirty] = useState(false);
  const [clusters, setClusters] = useState<ClusterDefinition[]>([]);
  const [activeSheet, setActiveSheet] = useState<SheetDefinition | null>(null);
  const [selectedClusterId, setSelectedClusterId] = useState<string | null>(
    null
  );
  const fileInputRef = useRef<HTMLInputElement>(null);
  const previewRef = useRef<HTMLDivElement>(null);

  const bgUrl = activeSheet?.backgroundImage
    ? `${API_BASE_URL}${activeSheet.backgroundImage}`
    : null;

  useEffect(() => {
    const unsub = controller.subscribe((state) => {
      setDirty(state.dirty);
      setClusters(state.form.clusters);
      setSelectedClusterId(state.selectedClusterId);
      const sheet = state.form.sheets.find(
        (s) => s.id === state.activeSheetId
      );
      setActiveSheet(sheet ?? null);
    });
    return unsub;
  }, []);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] ?? null;
    if (file) {
      setSelectedFile(file);
      setMessage(null);
      setIsError(false);
    }
  };

  const handleUpload = async (e: FormEvent) => {
    e.preventDefault();
    if (!selectedFile) {
      setMessage("Please select an Excel file first.");
      setIsError(true);
      return;
    }

    setIsLoading(true);
    setMessage(null);
    setIsError(false);

    try {
      const formData = new FormData();
      formData.append("file", selectedFile);

      const response = await fetch(`${API_BASE_URL}/api/form/from-excel`, {
        method: "POST",
        body: formData,
      });

      const result: ApiResponse<FormDefinition> = await response.json();

      if (response.ok && result.success && result.data) {
        controller.loadForm(result.data);
        const sheet = result.data.sheets[0];
        setZoom(0.5);
        if (sheet?.backgroundImage) {
          setMessage(
            `Imported ${result.data.clusters.length} cluster(s) — ${Math.round(sheet.backgroundWidth || 0)}×${Math.round(sheet.backgroundHeight || 0)}px background`
          );
        } else {
          setMessage(
            `${result.data.clusters.length} cluster(s) loaded from Excel.`
          );
        }
      } else {
        setMessage(result.message || "Failed to process file.");
        setIsError(true);
      }
    } catch (err) {
      setMessage(
        `Error: ${err instanceof Error ? err.message : "Unknown error"}`
      );
      setIsError(true);
    } finally {
      setIsLoading(false);
    }
  };

  const handleSave = async () => {
    setIsSaving(true);
    setMessage(null);
    setIsError(false);

    try {
      const form = controller.save();
      const response = await fetch(`${API_BASE_URL}/api/form/save`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(form),
      });

      const result: ApiResponse<{
        xmlPath: string;
        workbookPath: string;
        previewPath: string;
      }> = await response.json();

      if (response.ok && result.success) {
        setMessage(`Form saved successfully. ${result.message}`);
      } else {
        setMessage(result.message || "Failed to save form.");
        setIsError(true);
      }
    } catch (err) {
      setMessage(
        `Save error: ${err instanceof Error ? err.message : "Unknown error"}`
      );
      setIsError(true);
    } finally {
      setIsSaving(false);
    }
  };

  const handleReset = () => {
    setSelectedFile(null);
    setMessage(null);
    setIsError(false);
    setZoom(0.5);
    controller.resetForm();
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const handleClusterClick = (clusterId: string) => {
    if (selectedClusterId === clusterId) {
      controller.selectCluster(null);
    } else {
      controller.selectCluster(clusterId);
    }
  };

  const zoomIn = () => setZoom((z) => Math.min(z + 0.25, 3));
  const zoomOut = () => setZoom((z) => Math.max(z - 0.25, 0.25));
  const zoomReset = () => setZoom(1);

  const fieldTypeColor = (type: string) => {
    switch (type.toLowerCase()) {
      case "text":
        return "#3B82F6";
      case "date":
        return "#10B981";
      case "checkbox":
        return "#F59E0B";
      case "signature":
        return "#8B5CF6";
      case "number":
        return "#EF4444";
      case "barcode":
        return "#EC4899";
      case "qrcode":
        return "#6366F1";
      case "image":
        return "#14B8A6";
      case "table":
        return "#F97316";
      default:
        return "#FFD400";
    }
  };

  return (
    <div className="flex flex-col min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 dark:from-slate-950 dark:to-slate-900">
      <header className="border-b border-slate-200 dark:border-slate-800 bg-white/80 dark:bg-slate-900/80 backdrop-blur-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-4 flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center shadow-sm">
            <svg
              className="w-5 h-5 text-white"
              fill="none"
              viewBox="0 0 24 24"
              stroke="currentColor"
              strokeWidth={2}
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
              />
            </svg>
          </div>
          <div>
            <h1 className="text-lg font-semibold text-slate-900 dark:text-slate-50">
              FormLess Designer
            </h1>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              ConMas Desktop Designer &mdash; Form Definition Model
            </p>
          </div>
          <div className="ml-auto flex items-center gap-3">
            {dirty && (
              <span className="text-xs text-amber-500 font-medium flex items-center gap-1">
                <span className="w-2 h-2 rounded-full bg-amber-500 animate-pulse" />
                Unsaved changes
              </span>
            )}
            <button
              type="button"
              onClick={controller.undo.bind(controller)}
              disabled={!dirty}
              className="p-1.5 rounded-md hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-500 disabled:opacity-30"
              title="Undo"
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
                  d="M3 10h10a5 5 0 015 5v2M3 10l4-4m-4 4l4 4"
                />
              </svg>
            </button>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl mx-auto w-full px-4 sm:px-6 py-8 sm:py-12">
        <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-800 overflow-hidden">
          <div className="p-6 sm:p-8">
            <h2 className="text-xl font-semibold text-slate-900 dark:text-slate-50 mb-2">
              Import Excel File
            </h2>
            <p className="text-sm text-slate-500 dark:text-slate-400 mb-6">
              Import an Excel workbook to create a Form Definition. The
              spreadsheet&apos;s print area is rendered as a PNG background
              image. Cell comments define clusters (form fields) positioned as
              overlays.
            </p>

            <form onSubmit={handleUpload} className="space-y-5">
              <div className="relative">
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileChange}
                  className="block w-full text-sm text-slate-500 file:mr-4 file:py-2.5 file:px-5 file:rounded-lg file:border-0 file:text-sm file:font-medium file:bg-emerald-50 file:text-emerald-700 hover:file:bg-emerald-100 dark:file:bg-emerald-950 dark:file:text-emerald-300 dark:hover:file:bg-emerald-900/80 cursor-pointer transition-colors"
                />
                {selectedFile && (
                  <p className="mt-2 text-xs text-slate-400 dark:text-slate-500 flex items-center gap-1.5">
                    <svg
                      className="w-3.5 h-3.5 text-emerald-500"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={2}
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
                      />
                    </svg>
                    {selectedFile.name} ({(selectedFile.size / 1024).toFixed(
                      1
                    )}{" "}
                    KB)
                  </p>
                )}
              </div>

              <div className="flex gap-3">
                <button
                  type="submit"
                  disabled={!selectedFile || isLoading}
                  className="inline-flex items-center gap-2 px-6 py-2.5 rounded-lg bg-gradient-to-r from-emerald-600 to-teal-600 text-white text-sm font-medium hover:from-emerald-500 hover:to-teal-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all duration-200 shadow-sm hover:shadow-md active:scale-[0.98]"
                >
                  {isLoading ? (
                    <>
                      <svg
                        className="animate-spin h-4 w-4 text-white"
                        xmlns="http://www.w3.org/2000/svg"
                        fill="none"
                        viewBox="0 0 24 24"
                      >
                        <circle
                          className="opacity-25"
                          cx="12"
                          cy="12"
                          r="10"
                          stroke="currentColor"
                          strokeWidth="4"
                        />
                        <path
                          className="opacity-75"
                          fill="currentColor"
                          d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"
                        />
                      </svg>
                      Processing...
                    </>
                  ) : (
                    <>
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
                      Import
                    </>
                  )}
                </button>

                {dirty && (
                  <button
                    type="button"
                    onClick={handleSave}
                    disabled={isSaving}
                    className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-gradient-to-r from-blue-600 to-indigo-600 text-white text-sm font-medium hover:from-blue-500 hover:to-indigo-500 disabled:opacity-50 transition-all duration-200 shadow-sm hover:shadow-md active:scale-[0.98]"
                  >
                    {isSaving ? (
                      <>
                        <svg
                          className="animate-spin h-4 w-4"
                          xmlns="http://www.w3.org/2000/svg"
                          fill="none"
                          viewBox="0 0 24 24"
                        >
                          <circle
                            className="opacity-25"
                            cx="12"
                            cy="12"
                            r="10"
                            stroke="currentColor"
                            strokeWidth="4"
                          />
                          <path
                            className="opacity-75"
                            fill="currentColor"
                            d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"
                          />
                        </svg>
                        Saving...
                      </>
                    ) : (
                      <>
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
                            d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"
                          />
                        </svg>
                        Save
                      </>
                    )}
                  </button>
                )}

                {activeSheet && (
                  <button
                    type="button"
                    onClick={handleReset}
                    className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg border border-slate-300 dark:border-slate-700 text-slate-700 dark:text-slate-300 text-sm font-medium hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors"
                  >
                    New Form
                  </button>
                )}
              </div>
            </form>
          </div>
        </div>

        {message && (
          <div
            className={`mt-5 p-4 rounded-xl text-sm ${
              isError
                ? "bg-red-50 dark:bg-red-950/50 text-red-700 dark:text-red-300 border border-red-200 dark:border-red-900"
                : "bg-emerald-50 dark:bg-emerald-950/50 text-emerald-700 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-900"
            }`}
          >
            <div className="flex items-start gap-2.5">
              {isError ? (
                <svg
                  className="w-5 h-5 mt-0.5 shrink-0 text-red-500"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                  strokeWidth={2}
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
                  />
                </svg>
              ) : (
                <svg
                  className="w-5 h-5 mt-0.5 shrink-0 text-emerald-500"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                  strokeWidth={2}
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
                  />
                </svg>
              )}
              <span>{message}</span>
            </div>
          </div>
        )}

        {activeSheet && (
          <div className="mt-6 bg-white dark:bg-slate-900 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-800 overflow-hidden">
            <div className="flex items-center justify-between px-6 py-3 border-b border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950">
              <div className="flex items-center gap-2 text-sm text-slate-500 dark:text-slate-400">
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
                    d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z"
                  />
                </svg>
                {activeSheet.backgroundWidth > 0 && (
                  <span>
                    {activeSheet.backgroundWidth} &times;{" "}
                    {activeSheet.backgroundHeight} px
                  </span>
                )}
                <span className="text-slate-300 dark:text-slate-600">|</span>
                <span>
                  {activeSheet.name}
                  {activeSheet.printArea && (
                    <>
                      <span className="text-slate-300 dark:text-slate-600 mx-1">
                        |
                      </span>
                      Print: {activeSheet.pageSettings.widthPt} x{" "}
                      {activeSheet.pageSettings.heightPt} pt
                    </>
                  )}
                </span>
                <span className="text-slate-300 dark:text-slate-600">|</span>
                <span>
                  {clusters.length} cluster{clusters.length !== 1 ? "s" : ""}
                </span>
              </div>

              <div className="flex items-center gap-1">
                <button
                  onClick={zoomOut}
                  disabled={zoom <= 0.25}
                  className="p-1.5 rounded-md hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400 disabled:opacity-30 transition-colors"
                  title="Zoom out"
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
                      d="M20 12H4"
                    />
                  </svg>
                </button>
                <button
                  onClick={zoomReset}
                  className="px-2 py-1 text-xs font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-800 rounded-md transition-colors min-w-[3rem] text-center"
                >
                  {Math.round(zoom * 100)}%
                </button>
                <button
                  onClick={zoomIn}
                  disabled={zoom >= 3}
                  className="p-1.5 rounded-md hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400 disabled:opacity-30 transition-colors"
                  title="Zoom in"
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
                      d="M12 4v16m8-8H4"
                    />
                  </svg>
                </button>
              </div>
            </div>

            <div className="overflow-auto max-h-[80vh] bg-slate-100 dark:bg-slate-950">
              <div
                ref={previewRef}
                className="relative inline-block"
                style={
                  activeSheet.backgroundWidth > 0
                    ? {
                        width: activeSheet.backgroundWidth,
                        minHeight: activeSheet.backgroundHeight,
                        transform: `scale(${zoom})`,
                        transformOrigin: "top left",
                      }
                    : { minHeight: 400 }
                }
              >
                {/* Background PNG — the visual canvas */}
                {bgUrl && activeSheet.backgroundWidth > 0 && (
                  <Image
                    src={bgUrl}
                    alt={`Sheet: ${activeSheet.name}`}
                    width={activeSheet.backgroundWidth}
                    height={activeSheet.backgroundHeight}
                    className="block"
                    unoptimized
                    priority
                    draggable={false}
                  />
                )}

                {/* Background-less fallback */}
                {!bgUrl && (
                  <div className="p-8 text-center text-slate-400 min-w-[400px]">
                    <svg
                      className="w-12 h-12 mx-auto mb-3 opacity-50"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={1}
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M9 17h6m-6-4h6m-6-4h6M5 21h14a2 2 0 002-2V5a2 2 0 00-2-2H5a2 2 0 00-2 2v14a2 2 0 002 2z"
                      />
                    </svg>
                    <p className="text-sm">
                      {clusters.length} cluster(s) loaded.
                    </p>
                    <p className="text-xs mt-1">
                      No background image. Save to generate XML and XLSX.
                    </p>
                  </div>
                )}

                {/* Cluster overlays — positioned relative to background */}
                {clusters.map((cluster) => {
                  const isSelected = selectedClusterId === cluster.clusterId;
                  const color = fieldTypeColor(cluster.type);
                  return (
                    <div
                      key={cluster.clusterId}
                      className={`absolute cursor-pointer group ${
                        isSelected ? "z-10" : "z-0"
                      }`}
                      style={{
                        left: cluster.left,
                        top: cluster.top,
                        width: Math.max(cluster.right - cluster.left, 4),
                        height: Math.max(cluster.bottom - cluster.top, 4),
                      }}
                      onClick={() => handleClusterClick(cluster.clusterId)}
                    >
                      {/* Selection highlight */}
                      <div
                        className={`absolute inset-0 rounded pointer-events-none transition-all duration-150 ${
                          isSelected
                            ? "ring-2 ring-blue-500 shadow-lg"
                            : "opacity-0 group-hover:opacity-100"
                        }`}
                        style={{
                          backgroundColor: `${color}15`,
                          border: `2px solid ${color}`,
                          borderRadius: "2px",
                        }}
                      />

                      {/* Default visible overlay with muted appearance */}
                      <div
                        className="absolute inset-0 rounded pointer-events-none"
                        style={{
                          backgroundColor: `${color}08`,
                          border: `1px solid ${color}40`,
                          borderRadius: "2px",
                        }}
                      />

                      {/* Cluster type badge */}
                      <div
                        className={`absolute -top-5 left-0 px-1.5 py-0.5 rounded text-[9px] font-medium whitespace-nowrap pointer-events-none transition-opacity duration-150 ${
                          isSelected
                            ? "opacity-100"
                            : "opacity-0 group-hover:opacity-100"
                        }`}
                        style={{ backgroundColor: color, color: "#fff" }}
                      >
                        {cluster.name || cluster.cellAddress} &middot;{" "}
                        {cluster.type}
                      </div>

                      {/* Selected: render resize handle */}
                      {isSelected && (
                        <div className="absolute -bottom-1.5 -right-1.5 w-3 h-3 bg-blue-500 border-2 border-white rounded-sm shadow cursor-se-resize z-20" />
                      )}
                    </div>
                  );
                })}

                {/* Guide lines around selected cluster */}
                {selectedClusterId &&
                  (() => {
                    const c = clusters.find(
                      (cl) => cl.clusterId === selectedClusterId
                    );
                    if (!c) return null;
                    return (
                      <>
                        <div
                          className="absolute top-0 w-px bg-blue-400/50 pointer-events-none"
                          style={{ left: c.left, height: "100vh" }}
                        />
                        <div
                          className="absolute left-0 h-px bg-blue-400/50 pointer-events-none"
                          style={{ top: c.top, width: "100vw" }}
                        />
                        <div
                          className="absolute top-0 w-px bg-blue-400/50 pointer-events-none"
                          style={{ left: c.right, height: "100vh" }}
                        />
                        <div
                          className="absolute left-0 h-px bg-blue-400/50 pointer-events-none"
                          style={{ top: c.bottom, width: "100vw" }}
                        />
                      </>
                    );
                  })()}
              </div>
            </div>

            {/* Legend bar */}
            {clusters.length > 0 && (
              <div className="px-6 py-3 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950">
                <div className="flex flex-wrap gap-3">
                  <span className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wide">
                    Clusters:
                  </span>
                  {clusters.map((cluster) => {
                    const color = fieldTypeColor(cluster.type);
                    const isSelected =
                      selectedClusterId === cluster.clusterId;
                    return (
                      <button
                        key={cluster.clusterId}
                        onClick={() => handleClusterClick(cluster.clusterId)}
                        className={`inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-xs transition-all ${
                          isSelected ? "ring-2 ring-blue-400" : ""
                        }`}
                        style={{
                          backgroundColor: `${color}15`,
                          color: color,
                          border: `1px solid ${color}40`,
                        }}
                      >
                        <span className="font-mono text-[10px] opacity-70">
                          {cluster.cellAddress}
                        </span>
                        <span>{cluster.type}</span>
                      </button>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}
      </main>

      <footer className="border-t border-slate-200 dark:border-slate-800 bg-white/80 dark:bg-slate-900/80 backdrop-blur-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-4 text-center text-xs text-slate-400 dark:text-slate-600">
          FormLess Designer &mdash; Background PNG + Cluster Overlays
        </div>
      </footer>
    </div>
  );
}
