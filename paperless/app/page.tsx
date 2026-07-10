"use client";

import { useState } from "react";
import { useRuntime } from "@/hooks/useRuntime";
import { FormPage } from "@/components/FormViewer/FormPage";
import type { RuntimeField } from "@/types/runtime";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:5090";

export default function Home() {
  const [templateId, setTemplateId] = useState("");
  const [zoom, setZoom] = useState(0.5);
  const [focusedField, setFocusedField] = useState<RuntimeField | null>(null);
  const [debug, setDebug] = useState(false);
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [uploadResult, setUploadResult] = useState<{
    templateId: string;
    fileName: string;
    previewUrl: string | null;
  } | null>(null);

  const {
    runtimeForm,
    loading: runtimeLoading,
    error: runtimeError,
    reload: reloadRuntime,
  } = useRuntime(templateId || null);

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
    setUploadResult(null);
    setFocusedField(null);

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

      if (result.templateId) {
        const tid = result.templateId;
        const fileName = uploadFile.name;
        const previewUrl = result.previewUrl || null;
        setUploadResult({ templateId: tid, fileName, previewUrl });
        setTemplateId(tid);
      } else {
        throw new Error(result.message || "Upload succeeded but no template ID was returned.");
      }
    } catch (err) {
      setUploadError(
        `Upload failed: ${err instanceof Error ? err.message : "Unknown error"}`
      );
    } finally {
      setUploading(false);
    }
  };

  const handleReset = () => {
    setUploadFile(null);
    setUploading(false);
    setUploadError(null);
    setUploadResult(null);
    setTemplateId("");
    setFocusedField(null);
    const fileInput = document.querySelector<HTMLInputElement>('input[type="file"]');
    if (fileInput) fileInput.value = "";
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
              FormLess
            </h1>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              Excel-powered Runtime Form Viewer
            </p>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl mx-auto w-full px-4 sm:px-6 py-8 sm:py-12">

        {/* ── Upload Card ────────────────────────────────────────────── */}
        <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-800 overflow-hidden">
          <div className="p-6 sm:p-8">
            <h2 className="text-xl font-semibold text-slate-900 dark:text-slate-50 mb-2">
              Upload Excel Workbook
            </h2>
            <p className="text-sm text-slate-500 dark:text-slate-400 mb-6">
              Upload an Excel workbook. The print area is rendered as a PNG background
              with interactive input fields overlaid at each detected form field.
              Excel is the designer — no manual configuration needed.
            </p>

            <div className="space-y-4">
              {/* File chooser */}
              <div className="relative">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(e) => {
                    const file = e.target.files?.[0] ?? null;
                    if (file) {
                      setUploadFile(file);
                      setUploadError(null);
                    }
                  }}
                  className="block w-full text-sm text-slate-500 file:mr-4 file:py-2.5 file:px-5 file:rounded-lg file:border-0 file:text-sm file:font-medium file:bg-emerald-50 file:text-emerald-700 hover:file:bg-emerald-100 dark:file:bg-emerald-950 dark:file:text-emerald-300 dark:hover:file:bg-emerald-900/80 cursor-pointer transition-colors"
                />
                {uploadFile && (
                  <p className="mt-2 text-xs text-slate-400 dark:text-slate-500 flex items-center gap-1.5">
                    <svg className="w-3.5 h-3.5 text-emerald-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                    {uploadFile.name} ({(uploadFile.size / 1024).toFixed(1)} KB)
                  </p>
                )}
              </div>

              {/* Upload + new buttons */}
              <div className="flex gap-3">
                <button
                  onClick={handleUpload}
                  disabled={!uploadFile || uploading}
                  className="inline-flex items-center gap-2 px-6 py-2.5 rounded-lg bg-gradient-to-r from-emerald-600 to-teal-600 text-white text-sm font-medium hover:from-emerald-500 hover:to-teal-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all duration-200 shadow-sm hover:shadow-md active:scale-[0.98]"
                >
                  {uploading ? (
                    <>
                      <svg className="animate-spin h-4 w-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                        <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                      </svg>
                      Uploading...
                    </>
                  ) : (
                    <>
                      <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                        <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                      </svg>
                      Upload &amp; View
                    </>
                  )}
                </button>

                {uploadResult && (
                  <button
                    type="button"
                    onClick={handleReset}
                    className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg border border-slate-300 dark:border-slate-700 text-slate-700 dark:text-slate-300 text-sm font-medium hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors"
                  >
                    <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                    </svg>
                    New Form
                  </button>
                )}
              </div>
            </div>

            {/* ── Advanced: manual template ID ───────────────────── */}
            <details className="mt-6 group">
              <summary className="text-xs text-slate-400 dark:text-slate-500 hover:text-slate-600 dark:hover:text-slate-300 cursor-pointer select-none transition-colors">
                <span className="inline-flex items-center gap-1.5">
                  <svg className="w-3.5 h-3.5 transition-transform group-open:rotate-90" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M9 5l7 7-7 7" />
                  </svg>
                  Advanced — Manual Template ID
                </span>
              </summary>
              <div className="mt-3 flex gap-3">
                <input
                  type="text"
                  value={templateId}
                  onChange={(e) => setTemplateId(e.target.value)}
                  placeholder="Paste template ID (e.g., a1b2c3d4...)"
                  className="flex-1 px-4 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-sm text-slate-900 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent transition-all"
                />
                <button
                  onClick={() => reloadRuntime()}
                  disabled={!templateId || runtimeLoading}
                  className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300 text-sm font-medium hover:bg-slate-300 dark:hover:bg-slate-600 disabled:opacity-50 transition-colors"
                >
                  {runtimeLoading ? (
                    <>
                      <svg className="animate-spin h-4 w-4" fill="none" viewBox="0 0 24 24">
                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                        <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                      </svg>
                      Loading...
                    </>
                  ) : (
                    'Load'
                  )}
                </button>
              </div>
            </details>
          </div>
        </div>

        {/* ── Upload completed info ──────────────────────────────── */}
        {uploadResult && runtimeForm && (
          <div className="mt-5 p-4 rounded-xl text-sm bg-emerald-50 dark:bg-emerald-950/50 text-emerald-700 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-900">
            <div className="flex items-start gap-2.5">
              <svg className="w-5 h-5 mt-0.5 shrink-0 text-emerald-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <div className="flex flex-wrap items-center gap-x-3 gap-y-1">
                <span className="font-medium">{uploadResult.fileName}</span>
                <span className="text-emerald-400 dark:text-emerald-500">|</span>
                <span>Sheets: {runtimeForm.sheets.length}</span>
                <span className="text-emerald-400 dark:text-emerald-500">|</span>
                <span>Fields: {runtimeForm.sheets.reduce((sum, s) => sum + s.fields.length, 0)}</span>
                <span className="text-emerald-400 dark:text-emerald-500">|</span>
                <span>{runtimeForm.dpi} DPI</span>
                <span className="text-emerald-400 dark:text-emerald-500">|</span>
                <span className="font-mono text-xs">ID: {uploadResult.templateId}</span>
              </div>
            </div>
          </div>
        )}

        {/* ── Runtime error ──────────────────────────────────────── */}
        {runtimeError && !uploading && (
          <div className="mt-5 p-4 rounded-xl text-sm bg-red-50 dark:bg-red-950/50 text-red-700 dark:text-red-300 border border-red-200 dark:border-red-900">
            <div className="flex items-start gap-2.5">
              <svg className="w-5 h-5 mt-0.5 shrink-0 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <span>{runtimeError}</span>
            </div>
          </div>
        )}

        {/* ── Upload error ───────────────────────────────────────── */}
        {uploadError && (
          <div className="mt-5 p-4 rounded-xl text-sm bg-red-50 dark:bg-red-950/50 text-red-700 dark:text-red-300 border border-red-200 dark:border-red-900">
            <div className="flex items-start gap-2.5">
              <svg className="w-5 h-5 mt-0.5 shrink-0 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <span>{uploadError}</span>
            </div>
          </div>
        )}

        {/* ── Runtime form viewer ────────────────────────────────── */}
        {runtimeForm && runtimeForm.sheets.length > 0 && (
          <div className="mt-6 bg-white dark:bg-slate-900 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-800 overflow-hidden">
            <div className="flex items-center justify-between px-6 py-3 border-b border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950">
              <div className="flex items-center gap-2 text-sm text-slate-500 dark:text-slate-400">
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                <span>{runtimeForm.workbookName}</span>
                <span className="text-slate-300 dark:text-slate-600">|</span>
                <span>{runtimeForm.sheets.length} sheet(s)</span>
                <span className="text-slate-300 dark:text-slate-600">|</span>
                <span>{runtimeForm.sheets.reduce((sum, s) => sum + s.fields.length, 0)} field(s)</span>
                <span className="text-slate-300 dark:text-slate-600">|</span>
                <span>{runtimeForm.dpi} DPI</span>
              </div>

              <div className="flex items-center gap-1">
                <button
                  onClick={() => setZoom((z) => Math.max(z - 0.25, 0.25))}
                  disabled={zoom <= 0.25}
                  className="p-1.5 rounded-md hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400 disabled:opacity-30"
                >
                  <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M20 12H4" />
                  </svg>
                </button>
                <button
                  onClick={() => setZoom(1)}
                  className="px-2 py-1 text-xs font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-800 rounded-md min-w-[3rem] text-center"
                >
                  {Math.round(zoom * 100)}%
                </button>
                <button
                  onClick={() => setZoom((z) => Math.min(z + 0.25, 3))}
                  disabled={zoom >= 3}
                  className="p-1.5 rounded-md hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400 disabled:opacity-30"
                >
                  <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4" />
                  </svg>
                </button>
                <div className="w-px h-5 bg-slate-300 dark:bg-slate-600 mx-1" />
                <button
                  onClick={() => setDebug((d) => !d)}
                  className={`px-2 py-1 text-xs font-medium rounded-md transition-colors ${
                    debug
                      ? "bg-green-100 text-green-700 dark:bg-green-900 dark:text-green-300"
                      : "text-slate-500 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-800"
                  }`}
                  title="Toggle overlay debug borders"
                >
                  Debug
                </button>
              </div>
            </div>

            <div className="overflow-auto max-h-[80vh] bg-slate-100 dark:bg-slate-950 p-4">
              {runtimeForm.sheets.map((sheet) => (
                <FormPage
                  key={sheet.index}
                  sheet={sheet}
                  templateId={templateId}
                  previewUrl={uploadResult?.previewUrl}
                  zoom={zoom}
                  onFieldFocus={(field) => setFocusedField(field)}
                  debug={debug}
                />
              ))}
            </div>

            {/* Focused field info */}
            {focusedField && (
              <div className="px-6 py-3 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950">
                <div className="flex items-center gap-3 text-xs text-slate-500 dark:text-slate-400">
                  <span className="font-mono font-medium text-slate-700 dark:text-slate-300">
                    {focusedField.cellReference}
                  </span>
                  <span className="px-1.5 py-0.5 rounded bg-slate-200 dark:bg-slate-800">
                    {focusedField.dataType}
                  </span>
                  {focusedField.required && (
                    <span className="text-red-500 font-medium">Required</span>
                  )}
                  {focusedField.readOnly && (
                    <span className="text-amber-500">Read-only</span>
                  )}
                  <span className="text-slate-300 dark:text-slate-600">
                    {focusedField.leftPx.toFixed(0)}×{focusedField.topPx.toFixed(0)}px
                    {' '}{focusedField.widthPx.toFixed(0)}×{focusedField.heightPx.toFixed(0)}
                  </span>
                </div>
              </div>
            )}
          </div>
        )}
      </main>

      <footer className="border-t border-slate-200 dark:border-slate-800 bg-white/80 dark:bg-slate-900/80 backdrop-blur-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-4 text-center text-xs text-slate-400 dark:text-slate-600">
          FormLess Runtime &mdash; Excel-powered Form Viewer
        </div>
      </footer>
    </div>
  );
}
