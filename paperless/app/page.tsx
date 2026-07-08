"use client";

import { useState, useRef, type FormEvent } from "react";
import Image from "next/image";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:5090";

interface ExcelField {
  id: string;
  cell: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  comment: string;
}

interface PageInfo {
  width: number;
  height: number;
}

interface CaptureResult {
  imageUrl?: string;
  page?: PageInfo;
  fields?: ExcelField[];
}

interface UploadResponse {
  success: boolean;
  message: string;
  data?: CaptureResult;
}

export default function Home() {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [captureResult, setCaptureResult] = useState<CaptureResult | null>(null);
  const [message, setMessage] = useState<string | null>(null);
  const [isError, setIsError] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [zoom, setZoom] = useState(1);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const previewRef = useRef<HTMLDivElement>(null);

  const imageUrl = captureResult?.imageUrl
    ? `${API_BASE_URL}${captureResult.imageUrl}`
    : null;

  const fields = captureResult?.fields ?? [];
  const page = captureResult?.page;

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] ?? null;
    setSelectedFile(file);
    setCaptureResult(null);
    setMessage(null);
    setIsError(false);
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
    setCaptureResult(null);

    try {
      const formData = new FormData();
      formData.append("file", selectedFile);

      const response = await fetch(`${API_BASE_URL}/api/excel/upload`, {
        method: "POST",
        body: formData,
      });

      const result: UploadResponse = await response.json();

      if (response.ok && result.success && result.data?.imageUrl) {
        setCaptureResult(result.data);
        setMessage(
          `${result.data.fields?.length ?? 0} field(s) detected.`,
        );
        setIsError(false);
        setZoom(0.5); // Default zoom to fit most screens
      } else {
        setMessage(result.message || "Failed to process file.");
        setIsError(true);
      }
    } catch (err) {
      setMessage(
        `Error connecting to server: ${err instanceof Error ? err.message : "Unknown error"}`,
      );
      setIsError(true);
    } finally {
      setIsLoading(false);
    }
  };

  const handleReset = () => {
    setSelectedFile(null);
    setCaptureResult(null);
    setMessage(null);
    setIsError(false);
    setZoom(1);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const zoomIn = () => setZoom((z) => Math.min(z + 0.25, 3));
  const zoomOut = () => setZoom((z) => Math.max(z - 0.25, 0.25));
  const zoomReset = () => setZoom(1);

  const fieldTypeColor = (type: string) => {
    switch (type.toLowerCase()) {
      case "text": return "#3B82F6";
      case "date": return "#10B981";
      case "checkbox": return "#F59E0B";
      case "signature": return "#8B5CF6";
      case "number": return "#EF4444";
      default: return "#FFD400";
    }
  };

  return (
    <div className="flex flex-col min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 dark:from-slate-950 dark:to-slate-900">
      {/* Header */}
      <header className="border-b border-slate-200 dark:border-slate-800 bg-white/80 dark:bg-slate-900/80 backdrop-blur-sm">
        <div className="max-w-5xl mx-auto px-4 sm:px-6 py-4 flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center shadow-sm">
            <svg className="w-5 h-5 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
          </div>
          <div>
            <h1 className="text-lg font-semibold text-slate-900 dark:text-slate-50">
              PaperLess Enterprise
            </h1>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              Excel Print Area Capture &mdash; Field Metadata
            </p>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 max-w-5xl mx-auto w-full px-4 sm:px-6 py-8 sm:py-12">
        {/* Upload Section */}
        <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-800 overflow-hidden">
          <div className="p-6 sm:p-8">
            <h2 className="text-xl font-semibold text-slate-900 dark:text-slate-50 mb-2">
              Upload Excel File
            </h2>
            <p className="text-sm text-slate-500 dark:text-slate-400 mb-6">
              Select an Excel file with a configured Print Area and field
              definitions in cell comments.
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
                    <svg className="w-3.5 h-3.5 text-emerald-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                    {selectedFile.name} ({(selectedFile.size / 1024).toFixed(1)} KB)
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
                      <svg className="animate-spin h-4 w-4 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
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
                      Upload & Capture
                    </>
                  )}
                </button>

                {captureResult && (
                  <button
                    type="button"
                    onClick={handleReset}
                    className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg border border-slate-300 dark:border-slate-700 text-slate-700 dark:text-slate-300 text-sm font-medium hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors"
                  >
                    Upload Another
                  </button>
                )}
              </div>
            </form>
          </div>
        </div>

        {/* Status Message */}
        {message && (
          <div className={`mt-5 p-4 rounded-xl text-sm ${
            isError
              ? "bg-red-50 dark:bg-red-950/50 text-red-700 dark:text-red-300 border border-red-200 dark:border-red-900"
              : "bg-emerald-50 dark:bg-emerald-950/50 text-emerald-700 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-900"
          }`}>
            <div className="flex items-start gap-2.5">
              {isError ? (
                <svg className="w-5 h-5 mt-0.5 shrink-0 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
              ) : (
                <svg className="w-5 h-5 mt-0.5 shrink-0 text-emerald-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
              )}
              <span>{message}</span>
            </div>
          </div>
        )}

        {/* Preview with Field Overlay */}
        {captureResult && imageUrl && page && (
          <div className="mt-6 bg-white dark:bg-slate-900 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-800 overflow-hidden">
            {/* Toolbar */}
            <div className="flex items-center justify-between px-6 py-3 border-b border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950">
              <div className="flex items-center gap-2 text-sm text-slate-500 dark:text-slate-400">
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z" />
                </svg>
                <span>{page.width} &times; {page.height} px</span>
                <span className="text-slate-300 dark:text-slate-600">|</span>
                <span>{fields.length} field(s)</span>
              </div>

              {/* Zoom Controls */}
              <div className="flex items-center gap-1">
                <button
                  onClick={zoomOut}
                  disabled={zoom <= 0.25}
                  className="p-1.5 rounded-md hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400 disabled:opacity-30 transition-colors"
                  title="Zoom out"
                >
                  <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M20 12H4" />
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
                  <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4" />
                  </svg>
                </button>
              </div>
            </div>

            {/* Canvas with PNG + Field Overlay */}
            {/*
              Zoom/pan approach:
              - The container div handles scrolling (pan).
              - The inner scaled div uses transform: scale(zoom) with transformOrigin: "top left".
              - Both the PNG and the field overlay share the same coordinate system (PNG pixels at 300 DPI).
              - The field overlay has position: absolute over the PNG.
              - Scaling via CSS transform keeps the PNG and overlay perfectly aligned.
            */}
            <div
              className="overflow-auto max-h-[80vh] bg-slate-100 dark:bg-slate-950"
              style={{ cursor: zoom > 1 ? "grab" : "default" }}
            >
              <div
                ref={previewRef}
                className="relative"
                style={{
                  width: page.width,
                  height: page.height,
                  transform: `scale(${zoom})`,
                  transformOrigin: "top left",
                }}
              >
                {/* PNG Image */}
                <Image
                  src={imageUrl}
                  alt="Excel Print Area Preview"
                  width={page.width}
                  height={page.height}
                  className="block"
                  unoptimized
                  priority
                  draggable={false}
                />

                {/* Field Overlay — fields are positioned in the same coordinate space as the PNG */}
                {fields.map((field) => {
                  const color = fieldTypeColor(field.type);
                  return (
                    <div
                      key={field.id}
                      className="absolute pointer-events-auto cursor-pointer group"
                      style={{
                        left: field.left,
                        top: field.top,
                        width: field.width,
                        height: field.height,
                        backgroundColor: `${color}20`,
                        border: `2px solid ${color}`,
                        borderRadius: "2px",
                        transition: "background-color 0.15s, box-shadow 0.15s",
                      }}
                      title={`${field.cell}: ${field.type}\n${field.comment}`}
                    >
                      {/* Field label shown on hover */}
                      <div
                        className="absolute -top-6 left-0 px-2 py-0.5 rounded text-[10px] font-medium whitespace-nowrap opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none"
                        style={{ backgroundColor: color, color: "#fff" }}
                      >
                        {field.cell} &middot; {field.type}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Field Legend */}
            {fields.length > 0 && (
              <div className="px-6 py-3 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950">
                <div className="flex flex-wrap gap-3">
                  <span className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wide">
                    Fields:
                  </span>
                  {fields.map((field) => {
                    const color = fieldTypeColor(field.type);
                    return (
                      <span
                        key={field.id}
                        className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-xs"
                        style={{
                          backgroundColor: `${color}15`,
                          color: color,
                          border: `1px solid ${color}40`,
                        }}
                      >
                        <span className="font-mono text-[10px] opacity-70">{field.cell}</span>
                        <span>{field.type}</span>
                      </span>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-200 dark:border-slate-800 bg-white/80 dark:bg-slate-900/80 backdrop-blur-sm">
        <div className="max-w-5xl mx-auto px-4 sm:px-6 py-4 text-center text-xs text-slate-400 dark:text-slate-600">
          PaperLess Enterprise &mdash; Phase 2 &middot; Field Metadata Pipeline
        </div>
      </footer>
    </div>
  );
}
