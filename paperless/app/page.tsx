"use client";

import { useState, useCallback, useRef } from "react";
import type { RuntimeForm, RuntimeField } from "@/types/runtime";
import { useRuntime } from "@/hooks/useRuntime";
import { useRuntimeState } from "@/components/Runtime";
import { PaperlessDesigner } from "@/components/PaperlessDesigner";
import type { FormDefinition, SheetDefinition, ClusterDefinition } from "@/types/formDefinition";

// ── WorkbookDefinition type for save-edited flow (Phase 5.2) ──
interface WbDefField {
  cell: { address: string; rowIndex: number };
  name: string;
  type: string;
  value: string | null;
}

interface WbDefSheet {
  name: string;
  index: number;
  fields: WbDefField[];
}

interface WbDef {
  info: { title: string };
  sourceFileName: string;
  sessionId: string;
  sheets: WbDefSheet[];
}

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:5090";

/**
 * Convert a FormDefinition (upload-excel response) into a RuntimeForm for the Designer.
 */
function formDefinitionToRuntimeForm(fd: FormDefinition, fileName: string): RuntimeForm {
  const ptToPx = 300 / 72; // 300 DPI → pixels from points

  return {
    workbookName: fileName,
    title: fileName,
    pageWidth: fd.sheets[0]?.backgroundWidth ?? 2550,
    pageHeight: fd.sheets[0]?.backgroundHeight ?? 3300,
    scale: 1,
    dpi: 300,
    version: fd.workbook?.version ?? "1.0",
    sheets: fd.sheets.map((s: SheetDefinition, sheetIdx: number) => {
      const pageW = s.backgroundWidth > 0 ? s.backgroundWidth : Math.round(s.pageSettings.widthPt * ptToPx);
      const pageH = s.backgroundHeight > 0 ? s.backgroundHeight : Math.round(s.pageSettings.heightPt * ptToPx);

      // Filter clusters belonging to this sheet
      const sheetClusters = fd.clusters.filter((c: ClusterDefinition) => c.sheetId === s.id);

      return {
        name: s.name,
        index: sheetIdx,
        pageWidthPx: pageW,
        pageHeightPx: pageH,
        backgroundImage: null, // Uploaded Excel workbooks don't have background images
        images: [],
        shapes: [],
        printArea: null,
        fields: sheetClusters.map((c: ClusterDefinition, fi: number) => {
          const leftPx = Math.round(c.leftPt * ptToPx);
          const topPx = Math.round(c.topPt * ptToPx);
          const widthPx = Math.round(c.widthPt * ptToPx);
          const heightPx = Math.round(c.heightPt * ptToPx);
          const isMerged = c.cellAddress.includes(":");
          // Read existing value from cellValues
          const key = c.cellAddress.includes(":") ? c.cellAddress.split(":")[0] : c.cellAddress;
          const existingValue = s.cellValues?.[key] ?? "";

          return {
            id: c.clusterId,
            name: c.name,
            cellReference: c.cellAddress,
            row: 0,
            column: 0,
            leftPx,
            topPx,
            widthPx,
            heightPx,
            leftRatio: pageW > 0 ? leftPx / pageW : 0,
            topRatio: pageH > 0 ? topPx / pageH : 0,
            widthRatio: pageW > 0 ? widthPx / pageW : 0,
            heightRatio: pageH > 0 ? heightPx / pageH : 0,
            mergeRange: isMerged ? c.cellAddress : null,
            isMerged,
            dataType: mapClusterType(c.type),
            readOnly: c.readonly,
            required: c.inputParameters?.["required"] === "true",
            alignment: null,
            font: null,
            fontSize: 0,
            bold: false,
            fontColor: null,
            backgroundColor: null,
            border: null,
            placeholder: null,
            defaultValue: existingValue || null,
            maxLength: 0,
            tabIndex: fi,
            config: {
              appearance: {},
              behavior: { readOnly: c.readonly },
              input: {},
              layout: {},
            },
          };
        }),
      };
    }),
  };
}

/** Map a cluster type from FormDefinition to RuntimeField dataType */
function mapClusterType(type: string): RuntimeField["dataType"] {
  const t = (type ?? "").toLowerCase();
  if (t.includes("keyboard") || t.includes("text") || t === "") return "KeyboardText";
  if (t.includes("number") || t.includes("numeric")) return "number";
  if (t.includes("date") || t.includes("calendar")) return "date";
  if (t.includes("check") || t.includes("bool")) return "checkbox";
  if (t.includes("sign")) return "signature";
  if (t.includes("dropdown") || t.includes("select") || t.includes("list")) return "dropdown";
  if (t.includes("calc") || t.includes("formula")) return "calculated";
  return "KeyboardText";
}

function mapPreviewType(
  type: string,
):
  | "number"
  | "date"
  | "checkbox"
  | "signature"
  | "dropdown"
  | "calculated"
  | "KeyboardText" {
  const t = type.toLowerCase();
  if (t.includes("numeric") || t.includes("number")) return "number";
  if (t.includes("date")) return "date";
  if (t.includes("checkbox") || t.includes("radio") || t.includes("bool"))
    return "checkbox";
  if (t.includes("signature")) return "signature";
  if (t.includes("dropdown") || t.includes("select") || t.includes("list"))
    return "dropdown";
  if (t.includes("calc") || t.includes("formula") || t.includes("computed"))
    return "calculated";
  return "KeyboardText";
}
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
    setRuntimeForm,
  } = useRuntime(null);

  // ── Runtime state — manages field values ──
  const runtime = useRuntimeState();

  // ── Session ID from upload — server owns the workbook, browser only tracks this ──
  const [sessionId, setSessionId] = useState<string | null>(null);

  // ── Export state ──
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);
  const [exportSuccess, setExportSuccess] = useState<string | null>(null);

  // ── Upload Excel state (for the round-trip upload-excel flow) ──
  const [uploadExcelLoading, setUploadExcelLoading] = useState(false);
  const [uploadExcelError, setUploadExcelError] = useState<string | null>(null);

  // ── UI state ──
  const [isDragOver, setIsDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const uploadExcelInputRef = useRef<HTMLInputElement>(null);

  // ── Upload handler ──
  const handleUpload = async () => {
    if (!uploadFile) {
      setUploadError("Please select an Excel file first.");
      return;
    }

    const ext = uploadFile.name.split(".").pop()?.toLowerCase();
    if (!ext || (ext !== "xlsx" && ext !== "xls")) {
      setUploadError(
        "Invalid file extension. Please select an .xlsx or .xls file.",
      );
      return;
    }

    setUploading(true);
    setUploadError(null);
    runtime.reset();

    try {
      const formData = new FormData();
      formData.append("file", uploadFile);

      const response = await fetch(`${API_BASE_URL}/api/form/upload-preview`, {
        method: "POST",
        body: formData,
      });

      const result = await response.json();

      if (!response.ok || !result.success) {
        throw new Error(
          result.message || `HTTP ${response.status}: ${response.statusText}`,
        );
      }

      setTemplateName(uploadFile.name);

      // Phase 5.2: Store sessionId — server owns the workbook
      if (result.sessionId) {
        setSessionId(result.sessionId);
      }

      // Convert preview response to RuntimeForm format (multi-page)
      const pageList = result.pages ?? [];

      const runtimeForm: RuntimeForm = {
        workbookName: uploadFile.name,
        title: uploadFile.name,
        pageWidth: pageList[0]?.page?.width ?? 2550,
        pageHeight: pageList[0]?.page?.height ?? 3300,
        scale: 1,
        dpi: 300,
        version: "1.0",
        sheets: pageList.map((p: any, sheetIdx: number) => {
          const pageW = p.page?.width ?? 2550;
          const pageH = p.page?.height ?? 3300;
          const bgUrl = p.backgroundImage ?? "";

          return {
            name: p.sheetName ?? `Page ${sheetIdx + 1}`,
            index: sheetIdx,
            pageWidthPx: pageW,
            pageHeightPx: pageH,
            backgroundImage: bgUrl,
            images: [],
            shapes: [],
            printArea: null,
            fields: (p.fields ?? []).map((f: any, i: number) => {
              const leftRatio = f.left_ratio ?? 0;
              const topRatio = f.top_ratio ?? 0;
              const rightRatio = f.right_ratio ?? 0;
              const bottomRatio = f.bottom_ratio ?? 0;
              return {
                id: f.id,
                name: f.name,
                cellReference: f.cellAddr ?? "",
                row: 0,
                column: 0,
                leftPx: leftRatio * pageW,
                topPx: topRatio * pageH,
                widthPx: (rightRatio - leftRatio) * pageW,
                heightPx: (bottomRatio - topRatio) * pageH,
                leftRatio,
                topRatio,
                widthRatio: rightRatio - leftRatio,
                heightRatio: bottomRatio - topRatio,
                mergeRange: (f.cellAddr ?? "").includes(":")
                  ? f.cellAddr
                  : null,
                isMerged: (f.cellAddr ?? "").includes(":"),
                dataType: mapPreviewType(f.type ?? ""),
                readOnly: false,
                required: false,
                alignment: null,
                font: null,
                fontSize: 0,
                bold: false,
                fontColor: null,
                backgroundColor: null,
                border: null,
                placeholder: null,
                defaultValue: null,
                maxLength: 0,
                tabIndex: i,
              };
            }),
          };
        }),
      };

      setRuntimeForm(runtimeForm);
    } catch (err) {
      setUploadError(
        `Upload failed: ${err instanceof Error ? err.message : "Unknown error"}`,
      );
    } finally {
      setUploading(false);
      setUploadFile(null);
    }
  };

  // ── Reset handler ──
  const handleReset = () => {
    setTemplateId(null);
    setTemplateName("");
    setUploadFile(null);
    setUploading(false);
    setUploadError(null);
    setSessionId(null);
    runtime.reset();
    setRuntimeForm(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  // ── Upload button from designer toolbar → go to upload screen + open file dialog ──
  const handleUploadClick = useCallback(() => {
    handleReset();
    setTimeout(() => fileInputRef.current?.click(), 100);
  }, []);

  // ── Upload Excel handler — POST /api/form/upload-excel → reconstruct form ──
  const handleUploadExcel = useCallback(async (file: File) => {
    setUploadExcelLoading(true);
    setUploadExcelError(null);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch(`${API_BASE_URL}/api/form/upload-excel`, {
        method: "POST",
        body: formData,
      });

      const result = await response.json();

      if (!response.ok || !result.success) {
        throw new Error(result.message || `HTTP ${response.status}: ${response.statusText}`);
      }

      const fd = result.data.formDefinition as FormDefinition;
      if (!fd || !fd.sheets || fd.sheets.length === 0) {
        throw new Error("No sheets found in the uploaded workbook.");
      }

      setTemplateName(file.name);
      runtime.reset();

      // Phase 5.2: Store sessionId — server owns the workbook
      if (result.data?.sessionId) {
        setSessionId(result.data.sessionId);
      }

      // Convert FormDefinition → RuntimeForm and set it
      const converted = formDefinitionToRuntimeForm(fd, file.name);

      // Pre-populate the runtime store with existing values from cellValues
      for (const sheet of fd.sheets) {
        for (const cluster of fd.clusters.filter(c => c.sheetId === sheet.id)) {
          const key = cluster.cellAddress.includes(":")
            ? cluster.cellAddress.split(":")[0]
            : cluster.cellAddress;
          const val = sheet.cellValues?.[key];
          if (val && val.trim()) {
            runtime.setValue(cluster.clusterId, val);
          }
        }
      }

      setRuntimeForm(converted);
      setExportSuccess(null);
      setExportError(null);
    } catch (err) {
      setUploadExcelError(err instanceof Error ? err.message : "Upload failed");
    } finally {
      setUploadExcelLoading(false);
      if (uploadExcelInputRef.current) {
        uploadExcelInputRef.current.value = "";
      }
    }
  }, [runtime, setRuntimeForm]);

  // ── Convert RuntimeForm → WorkbookDefinition for save-edited flow (Phase 5.2) ──
  const runtimeFormToWorkbookDefinition = useCallback((
    form: RuntimeForm,
    values: Record<string, string | boolean | null>,
    sid: string,
  ): WbDef => {
    return {
      info: { title: form.workbookName ?? "Untitled" },
      sourceFileName: sid,
      sessionId: sid,
      sheets: form.sheets.map((sheet, si) => ({
        name: sheet.name ?? `Page ${si + 1}`,
        index: si,
        fields: sheet.fields
          .filter(f => {
            const val = values[f.id];
            return val !== null && val !== undefined && val !== "";
          })
          .map(f => ({
            cell: {
              address: f.cellReference?.split(":")[0] ?? "A1",
              rowIndex: parseInt(f.cellReference?.match(/\d+/)?.[0] ?? "1"),
            },
            name: f.name ?? f.id,
            type: f.dataType ?? "KeyboardText",
            value: String(values[f.id] ?? ""),
          })),
      })),
    };
  }, []);

  // ── Save edited — POST /api/form/save-edited → WorkbookValueWriter (Phase 5.2) ──
  const handleSaveEdited = useCallback(async () => {
    if (!runtimeForm || !sessionId) {
      setExportError(sessionId ? "No form loaded." : "No session found. Please upload the workbook first.");
      return;
    }

    setExporting(true);
    setExportError(null);
    setExportSuccess(null);

    try {
      const wbDef = runtimeFormToWorkbookDefinition(runtimeForm, runtime.values, sessionId);

      const response = await fetch(`${API_BASE_URL}/api/form/save-edited`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(wbDef),
      });

      const contentType = response.headers.get("content-type") ?? "";

      if (!response.ok || contentType.includes("application/json")) {
        let errorMsg = `HTTP ${response.status}: ${response.statusText}`;
        try {
          const errResult = await response.clone().json();
          errorMsg = errResult.message ?? errorMsg;
        } catch {}
        throw new Error(errorMsg);
      }

      // Log validation results from response header
      const validationHeader = response.headers.get("X-Validation-Results");
      if (validationHeader) {
        try {
          const validation = JSON.parse(atob(validationHeader));
          console.log("[SAVE-EDITED] Workbook validation results:", validation);
        } catch {}
      }

      // Download the workbook
      const blob = await response.blob();
      const disposition = response.headers.get("content-disposition") ?? "";
      let fileName = `${runtimeForm.workbookName.replace(/\.\w+$/, "")}_edited.xlsx`;
      const match = disposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
      if (match && match[1]) {
        fileName = match[1].replace(/['"]/g, "");
      }

      const blobUrl = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = blobUrl;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(blobUrl);

      setExportSuccess("Workbook saved successfully! Values written into original workbook.");
      runtime.markAllClean();
      setTimeout(() => setExportSuccess(null), 3000);
    } catch (err) {
      setExportError(err instanceof Error ? err.message : "Save failed");
    } finally {
      setExporting(false);
    }
  }, [runtimeForm, runtime.values, sessionId, runtimeFormToWorkbookDefinition, runtime]);

  // ── Export Excel handler — Phase 5.2: save-edited is the ONLY path ──
  const handleExportExcel = useCallback(async () => {
    if (!runtimeForm) {
      setExportError("No form loaded.");
      return;
    }

    if (!sessionId) {
      setExportError("No session found. Please upload the workbook first.");
      return;
    }

    return handleSaveEdited();
  }, [runtimeForm, sessionId, handleSaveEdited]);

  // ── Hidden file input for the Upload Excel button ──
  const handleUploadExcelFileSelected = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      handleUploadExcel(file);
    }
  }, [handleUploadExcel]);

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
        <div className="mx-auto px-3  py-3 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center shadow-sm">
              <svg
                className="w-4 h-4 text-white"
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
                  d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"
                />
              </svg>
              Open Template
            </button>

            {hasTemplate && (
              <>
                {runtime.isDirty() && (
                  <button
                    onClick={handleReset}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-600 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors"
                    title="Reset all form fields"
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
                        d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"
                      />
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
      <main className="flex-1 w-full overflow-hidden ">
        {!hasTemplate ? (
          /* ── Upload Screen ───────────────────────────────── */
          <div className="flex flex-col items-center justify-center pt-12 sm:pt-20">
            {/* Logo / Branding */}
            <div className="mb-8 text-center">
              <div className="w-16 h-16 rounded-2xl bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center shadow-lg mx-auto mb-4">
                <svg
                  className="w-8 h-8 text-white"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                  strokeWidth={1.5}
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z"
                  />
                </svg>
              </div>
              <h2 className="text-2xl sm:text-3xl font-bold text-slate-900 mb-2">
                PaperLess Enterprise
              </h2>
              <p className="text-sm text-slate-500 max-w-md mx-auto">
                Upload an Excel template. Excel renders the form. Fill it in the
                browser.
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
                ${
                  isDragOver
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
                <div
                  className={`w-12 h-12 rounded-full flex items-center justify-center transition-colors ${
                    isDragOver ? "bg-emerald-100" : "bg-slate-100"
                  }`}
                >
                  <svg
                    className={`w-6 h-6 transition-colors ${
                      isDragOver ? "text-emerald-600" : "text-slate-400"
                    }`}
                    fill="none"
                    viewBox="0 0 24 24"
                    stroke="currentColor"
                    strokeWidth={1.5}
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"
                    />
                  </svg>
                </div>

                <div>
                  <p className="text-sm font-medium text-slate-700">
                    {isDragOver
                      ? "Drop your file here"
                      : "Choose an Excel file"}
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
                  <svg
                    className="w-4 h-4 text-emerald-500"
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
                  <span className="text-sm text-slate-700">
                    {uploadFile.name}
                  </span>
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
                      <svg
                        className="animate-spin h-4 w-4"
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
                  <svg
                    className="w-4 h-4 mt-0.5 shrink-0 text-red-500"
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
                  <span>{uploadError}</span>
                </div>
              </div>
            )}

            {/* Runtime error */}
            {runtimeError && (
              <div className="mt-4 p-3 rounded-xl text-sm bg-red-50 text-red-700 border border-red-200 max-w-lg">
                <div className="flex items-start gap-2">
                  <svg
                    className="w-4 h-4 mt-0.5 shrink-0 text-red-500"
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
                  <span>{runtimeError}</span>
                </div>
              </div>
            )}
          </div>
        ) : (
          /* ── Designer View ───────────────────────────────── */
          <div className=" flex flex-col h-[100vh] ">
            <PaperlessDesigner
              runtimeForm={runtimeForm}
              runtime={runtime}
              templateName={templateName}
              onReset={handleReset}
              onUploadClick={handleUploadClick}
              onUploadExcel={() => uploadExcelInputRef.current?.click()}
              onExportExcel={handleExportExcel}
              exporting={exporting}
              exportError={exportError}
              exportSuccess={exportSuccess}
            />
            {/* Hidden file input for Upload Excel */}
            <input
              ref={uploadExcelInputRef}
              type="file"
              accept=".xlsx"
              className="hidden"
              onChange={handleUploadExcelFileSelected}
            />
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
