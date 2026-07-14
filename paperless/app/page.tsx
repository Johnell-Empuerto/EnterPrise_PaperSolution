"use client";

import { useState, useCallback, useRef } from "react";
import { useRuntime } from "@/hooks/useRuntime";
import { RuntimeFormViewer, useRuntimeState } from "@/components/Runtime";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:5090";

export default function Home() {
  // ── Upload state ──
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);

  // ── Template ID → fetch RuntimeForm via COM backend ──
  const [templateId, setTemplateId] = useState<string | null>(null);
  const [templateName, setTemplateName] = useState<string>("");

  const {
    runtimeForm,
    loading: runtimeLoading,
    error: runtimeError,
    loadByTemplateId,
  } = useRuntime(null);

  // ── Runtime state — manages field values ──
  const runtime = useRuntimeState();

  // ── UI state ──
  const [isDragOver, setIsDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── Upload handler ──
  const handleUpload = async () => {
    if (!uploadFile) {
      setUploadError("Please select an Excel file first.");
      return;
    }

    const ext = uploadFile.name.split(".").pop()?.toLowerCase();
    if (!ext || (ext !== "xlsx" && ext !== "xls")) {
      setUploadError("Invalid file extension. Please select an .xlsx or .xls file.");
      return;
    }

    setUploading(true);
    setUploadError(null);
    runtime.reset();

    try {
      const formData = new FormData();
      formData.append("file", uploadFile);

      const response = await fetch(`${API_BASE_URL}/api/form/from-excel`, {
        method: "POST",
        body: formData,
      });

      const result = await response.json();

      if (!response.ok || !result.success) {
        throw new Error(result.message || `HTTP ${response.status}: ${response.statusText}`);
      }

      if (!result.templateId) {
        throw new Error(result.message || "Upload succeeded but no template ID was returned.");
      }

      const tid = result.templateId;
      setTemplateId(tid);
      setTemplateName(uploadFile.name);

      // Fetch the RuntimeForm from the COM backend (includes PNG URL + pixel coordinates)
      await loadByTemplateId(tid);
    } catch (err) {
      setUploadError(
        `Upload failed: ${err instanceof Error ? err.message : "Unknown error"}`
      );
    } finally {
      setUploading(false);
      setUploadFile(null);
    }
  };

  // ── Print handler ──
  const handlePrint = useCallback(() => {
    window.print();
  }, []);

  // ── Reset handler ──
  const handleReset = () => {
    setTemplateId(null);
    setTemplateName("");
    setUploadFile(null);
    setUploading(false);
    setUploadError(null);
    runtime.reset();
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  // ── Drag and drop handlers ──
  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);

    const files = e.dataTransfer.files;
    if (files.length > 0) {
      const file = files[0];
      const ext = file.name.split(".").pop()?.toLowerCase();
      if (ext === "xlsx" || ext === "xls") {
        setUploadFile(file);
        setUploadError(null);
      } else {
        setUploadError("Please drop an .xlsx or .xls file.");
      }
    }
  }, []);

  const hasTemplate = runtimeForm !== null;

  return (
    <div className="flex flex-col min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">
      {/* ── Header ─────────────────────────────────────────── */}
      <header className="border-b border-slate-200 bg-white/90 backdrop-blur-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-3 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center shadow-sm">
              <svg className="w-4 h-4 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
            </div>
            <div>
              <h1 className="text-lg font-semibold text-slate-900">
                PaperLess
              </h1>
              <p className="text-[11px] text-slate-500">
                Intelligent Form Engine
              </p>
            </div>
          </div>

          {/* Production toolbar */}
          <div className="flex items-center gap-2 no-print">
            <button
              onClick={handleReset}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-600 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 hover:border-slate-400 transition-colors"
              title="Open a new template"
            >
              <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
              </svg>
              Open Template
            </button>

            {hasTemplate && (
              <>
                <button
                  onClick={handlePrint}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-white bg-emerald-600 rounded-lg hover:bg-emerald-500 transition-colors shadow-sm"
                  title="Print or Save as PDF"
                >
                  <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z" />
                  </svg>
                  Print / Save PDF
                </button>

                {runtime.isDirty() && (
                  <button
                    onClick={handleReset}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-600 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors"
                    title="Reset all form fields"
                  >
                    <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                    </svg>
                    Reset
                  </button>
                )}
              </>
            )}
          </div>
        </div>
      </header>

      {/* ── Main Content ───────────────────────────────────── */}
      <main className="flex-1 max-w-7xl mx-auto w-full px-4 sm:px-6 py-6 sm:py-10">
        {!hasTemplate ? (
          /* ── Upload Screen ───────────────────────────────── */
          <div className="flex flex-col items-center justify-center pt-12 sm:pt-20">
            {/* Logo / Branding */}
            <div className="mb-8 text-center">
              <div className="w-16 h-16 rounded-2xl bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center shadow-lg mx-auto mb-4">
                <svg className="w-8 h-8 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
                </svg>
              </div>
              <h2 className="text-2xl sm:text-3xl font-bold text-slate-900 mb-2">
                PaperLess Enterprise
              </h2>
              <p className="text-sm text-slate-500 max-w-md mx-auto">
                Upload an Excel template. Excel renders the form. Fill it in the browser.
              </p>
            </div>

            {/* Upload Area */}
            <div
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              className={`
                w-full max-w-lg cursor-pointer rounded-2xl border-2 border-dashed p-12 text-center
                transition-all duration-200
                ${isDragOver
                  ? "border-emerald-400 bg-emerald-50 scale-[1.02]"
                  : "border-slate-300 bg-white hover:border-emerald-300 hover:bg-emerald-50/50"
                }
              `}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0] ?? null;
                  if (file) {
                    setUploadFile(file);
                    setUploadError(null);
                  }
                }}
              />

              <div className="flex flex-col items-center gap-3">
                {/* Upload icon */}
                <div className={`w-12 h-12 rounded-full flex items-center justify-center transition-colors ${
                  isDragOver ? "bg-emerald-100" : "bg-slate-100"
                }`}>
                  <svg className={`w-6 h-6 transition-colors ${
                    isDragOver ? "text-emerald-600" : "text-slate-400"
                  }`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5" />
                  </svg>
                </div>

                <div>
                  <p className="text-sm font-medium text-slate-700">
                    {isDragOver ? "Drop your file here" : "Choose an Excel file"}
                  </p>
                  <p className="text-xs text-slate-400 mt-1">
                    Drag and drop or click to browse
                  </p>
                </div>

                <span className="text-[10px] text-slate-300 bg-slate-50 px-2 py-0.5 rounded">
                  .xlsx · .xls
                </span>
              </div>
            </div>

            {/* File selected indicator */}
            {uploadFile && (
              <div className="mt-4 flex items-center gap-3">
                <div className="flex items-center gap-2 px-4 py-2 bg-white rounded-lg border border-slate-200 shadow-sm">
                  <svg className="w-4 h-4 text-emerald-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                  <span className="text-sm text-slate-700">{uploadFile.name}</span>
                  <span className="text-xs text-slate-400">
                    ({(uploadFile.size / 1024).toFixed(1)} KB)
                  </span>
                </div>

                <button
                  onClick={handleUpload}
                  disabled={uploading || runtimeLoading}
                  className="inline-flex items-center gap-2 px-6 py-2.5 rounded-lg bg-gradient-to-r from-emerald-600 to-teal-600 text-white text-sm font-medium hover:from-emerald-500 hover:to-teal-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all duration-200 shadow-sm hover:shadow-md active:scale-[0.98]"
                >
                  {uploading ? (
                    <>
                      <svg className="animate-spin h-4 w-4" fill="none" viewBox="0 0 24 24">
                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                        <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                      </svg>
                      Processing...
                    </>
                  ) : (
                    <>
                      <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                        <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                      </svg>
                      Upload &amp; Open Form
                    </>
                  )}
                </button>
              </div>
            )}

            {/* Upload error */}
            {uploadError && (
              <div className="mt-4 p-3 rounded-xl text-sm bg-red-50 text-red-700 border border-red-200 max-w-lg">
                <div className="flex items-start gap-2">
                  <svg className="w-4 h-4 mt-0.5 shrink-0 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                  <span>{uploadError}</span>
                </div>
              </div>
            )}

            {/* Runtime error */}
            {runtimeError && (
              <div className="mt-4 p-3 rounded-xl text-sm bg-red-50 text-red-700 border border-red-200 max-w-lg">
                <div className="flex items-start gap-2">
                  <svg className="w-4 h-4 mt-0.5 shrink-0 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                  <span>{runtimeError}</span>
                </div>
              </div>
            )}

            {/* Developer link */}
            <div className="mt-12 text-center">
              <a
                href="/compare"
                className="text-xs text-slate-400 hover:text-slate-600 underline underline-offset-2 transition-colors"
              >
                Developer Tools (Debug / Compare)
              </a>
            </div>
          </div>
        ) : (
          /* ── Form View ──────────────────────────────────── */
          <div>
            {/* Template info bar */}
            <div className="mb-4 flex items-center justify-between">
              <div className="flex items-center gap-2 text-sm text-slate-500">
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                <span className="font-medium text-slate-700">{templateName}</span>
                <span className="text-slate-300">|</span>
                <span>{runtimeForm.sheets.length} sheet(s)</span>
                <span className="text-slate-300">|</span>
                <span>
                  {runtimeForm.sheets.reduce((sum, s) => sum + s.fields.length, 0)} field(s)
                </span>
                <span className="text-slate-300">|</span>
                <span>{runtimeForm.dpi} DPI</span>
              </div>
            </div>

            {/* PageSurface viewport — scrollable container acting as WPF ScrollViewer
                No padding, no flex, no justify-center — the PageSurface within is at
                fixed native pixel dimensions (e.g., 2550×3299) and must not be
                centered, scaled, or constrained. */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-auto">
              <RuntimeFormViewer
                runtimeForm={runtimeForm}
                runtime={runtime}
              />
            </div>

            {/* Dirty status indicator */}
            {runtime.isDirty() && (
              <div className="mt-3 flex items-center justify-between">
                <p className="text-xs text-amber-600 flex items-center gap-1.5">
                  <span className="w-1.5 h-1.5 rounded-full bg-amber-500 inline-block" />
                  Form has unsaved changes
                </p>
                <button
                  onClick={() => {
                    navigator.clipboard.writeText(runtime.exportJson());
                  }}
                  className="text-xs text-slate-400 hover:text-slate-600 underline underline-offset-2 transition-colors"
                >
                  Copy values as JSON
                </button>
              </div>
            )}
          </div>
        )}
      </main>

      {/* ── Footer ─────────────────────────────────────────── */}
      <footer className="border-t border-slate-200 bg-white/80 backdrop-blur-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-3 text-center text-[11px] text-slate-400">
          PaperLess Enterprise &mdash; Intelligent Form Engine
        </div>
      </footer>
    </div>
  );
}
