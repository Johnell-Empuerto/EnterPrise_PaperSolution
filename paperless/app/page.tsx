"use client";

import { useState, useCallback, useRef, useEffect } from "react";
import type { RuntimeForm, RuntimeField, FieldConfig } from "@/types/runtime";
import { useRuntime } from "@/hooks/useRuntime";
import { useRuntimeState } from "@/components/Runtime";
import { PaperlessDesigner } from "@/components/PaperlessDesigner";

// Phase 22: Style object matching backend CellStyle/FontDefinition/AlignmentDefinition.
// System.Text.Json camelCase serialization maps these to C# models directly.
interface WbDefStyle {
  font?: {
    name?: string;
    sizePt?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    colorArgb?: string;
  };
  alignment?: {
    horizontal?: string;
    vertical?: string;
  };
  fill?: {
    patternType?: string;
    colorArgb?: string;
  };
  wrapText?: boolean;
}

// ── WorkbookDefinition type for save-edited flow (Phase 5.2) ──
// PHASE 21.6: type is now a number (integer enum value matching backend FieldType)
interface WbDefField {
  id: string;
  cell: { address: string; rowIndex: number };
  name: string;
  type: number;
  value: string | null;
  // Phase 22: Browser style persistence — applied to cell after value write
  style?: WbDefStyle;
  // Full PaperLess config properties for round-trip persistence
  required?: boolean;
  locked?: boolean;
  visible?: boolean;
  maxLength?: number;
  placeholder?: string | null;
  defaultValue?: string | null;
  validateOnEditing?: boolean;
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

// ═════════════════════════════════════════════════════════
// PHASE 21.6 — Map frontend RuntimeField.dataType to backend
// FieldType integer enum value (no JsonStringEnumConverter
// registered on backend — System.Text.Json expects integers).
//
// Backend FieldType enum (in order):
//   Text=0, Number=1, Date=2, Checkbox=3,
//   Signature=4, Dropdown=5, Calculated=6
// ═════════════════════════════════════════════════════════
function fieldTypeToBackendEnum(dataType: string): number {
  switch (dataType) {
    case "number":     return 1; // Number
    case "date":       return 2; // Date
    case "checkbox":   return 3; // Checkbox
    case "signature":  return 4; // Signature
    case "dropdown":   return 5; // Dropdown
    case "calculated": return 6; // Calculated
    default:           return 0; // Text ("KeyboardText" + anything else)
  }
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

      // ═════════════════════════════════════════════════════════
      // PHASE 21.5 — STAGE 1: Upload Response → RuntimeForm
      // ═════════════════════════════════════════════════════════
      console.log("%c=========================================================", "color: #047857; font-weight: bold");
      console.log("%cSTAGE 1 — Upload Response → RuntimeForm", "color: #047857; font-weight: bold");
      console.log("%c=========================================================", "color: #047857; font-weight: bold");
      console.log("Upload Response:");
      console.log("  Pages:", pageList.length);
      pageList.forEach((p: any, pi: number) => {
        const pageFields = p.fields ?? [];
        console.log(`  Page ${pi}: '${p.sheetName ?? "?"}' — Fields: ${pageFields.length}`);
        pageFields.forEach((f: any, fi: number) => {
          console.log(`    Field ${fi}: id='${f.id ?? "?"}' name='${f.name ?? "?"}' cell='${f.cellAddr ?? "?"}' type='${f.type ?? "?"}'`);
        });
      });
      console.log("");
      console.log("Created RuntimeForm:");
      console.log("  workbookName:", runtimeForm.workbookName);
      console.log("  sheets:", runtimeForm.sheets.length);
      runtimeForm.sheets.forEach((s, si) => {
        console.log(`  Sheet ${si}: '${s.name}' — Fields: ${s.fields.length}`);
        s.fields.forEach((f, fi) => {
          console.log(`    Field ${fi}: id='${f.id}' name='${f.name ?? "?"}' cell='${f.cellReference}' type='${f.dataType}'`);
        });
      });
      console.log("%c=========================================================", "color: #047857; font-weight: bold");

            // ═════════════════════════════════════════════════════════
            // Apply PaperLessConfig to RuntimeForm fields
            // Restores field ID, font, size, bold, color, alignment,
            // required, maxLength, and type from the embedded config JSON.
            // Also builds runtimeField.config so KeyboardTextField and
            // KeyboardTextPropertyPanel reflect restored values.
            // ═════════════════════════════════════════════════════════
            function mapPLHorizontalAlignment(h: string | undefined): "left" | "center" | "right" | undefined {
              if (!h) return undefined;
              const a = h.toLowerCase();
              if (a === "left") return "left";
              if (a === "center" || a === "middle") return "center";
              if (a === "right") return "right";
              return undefined;
            }
            function mapPLVerticalAlignment(v: string | undefined): "top" | "middle" | "bottom" | undefined {
              if (!v) return undefined;
              const a = v.toLowerCase();
              if (a === "top") return "top";
              if (a === "center" || a === "middle") return "middle";
              if (a === "bottom") return "bottom";
              return undefined;
            }
            // Convert #AARRGGBB or #RRGGBB to #RRGGBB
            function argbToRgb(argb: string | undefined): string | undefined {
              if (!argb) return undefined;
              const c = argb.replace("#", "");
              if (c.length === 8) return "#" + c.substring(2);
              if (c.length === 6) return "#" + c.toUpperCase();
              if (c.length === 3) return "#" + c[0] + c[0] + c[1] + c[1] + c[2] + c[2];
              return argb;
            }
            if (result.paperLessConfig?.sheets) {
              for (const configSheet of result.paperLessConfig.sheets) {
                const runtimeSheet = runtimeForm.sheets.find(s => s.name === configSheet.name);
                if (!runtimeSheet) continue;

                for (const cf of configSheet.fields ?? []) {
                  const configCell = (cf.cell ?? "").replace(/^\$/, "").split(":")[0].replace(/\$/g, "").toUpperCase();
                  if (!configCell) continue;

                  const runtimeField = runtimeSheet.fields.find(f => {
                    const fCell = (f.cellReference ?? "").replace(/^\$/, "").split(":")[0].replace(/\$/g, "").toUpperCase();
                    return fCell === configCell;
                  });
                  if (!runtimeField) continue;

                  // Field identity
                  if (cf.id) runtimeField.id = cf.id;

                  // Style — font
                  const styleFont = cf.style?.font;
                  if (styleFont) {
                    if (styleFont.name) runtimeField.font = styleFont.name;
                    if (styleFont.sizePt > 0) runtimeField.fontSize = styleFont.sizePt;
                    if (styleFont.bold) runtimeField.bold = true;
                    if (styleFont.colorArgb) runtimeField.fontColor = styleFont.colorArgb;
                  }

                  // Style — fill / background
                  const styleFill = cf.style?.fill;
                  if (styleFill?.colorArgb) {
                    runtimeField.backgroundColor = styleFill.colorArgb;
                  }

                  // Style — alignment (horizontal + vertical)
                  const vAlignFromConfig = cf.style?.alignment?.vertical;
                  if (cf.style?.alignment?.horizontal) {
                    runtimeField.alignment = cf.style.alignment.horizontal;
                  }

                  // Field type
                  if (cf.type) {
                    runtimeField.dataType = mapPreviewType(cf.type);
                  }

                  // Input configuration (from PaperLessFieldConfig)
                  const cfg = cf.config;
                  if (cfg) {
                    if (cfg.required) runtimeField.required = true;
                    if (cfg.maxLength > 0) runtimeField.maxLength = cfg.maxLength;
                    if (cfg.placeholder !== undefined && cfg.placeholder !== null) runtimeField.placeholder = cfg.placeholder;
                    if (cfg.defaultValue !== undefined && cfg.defaultValue !== null) runtimeField.defaultValue = cfg.defaultValue;
                    // inputRestriction is stored in cfg but used via config.appearance below
                  }

                  // Build FieldConfig from PaperLessConfig so runtime components
                  // (KeyboardTextField, KeyboardTextPropertyPanel) reflect restored values.
                  // Without this, the property panel shows DEFAULTS (fontSize=11) and
                  // panel changes would override the correct flat property values.
                  runtimeField.config = {
                    appearance: {
                      fontFamily: styleFont?.name || undefined,
                      fontSize: (styleFont?.sizePt ?? 0) > 0 ? styleFont!.sizePt : undefined,
                      fontWeight: styleFont?.bold ? "bold" : undefined,
                      textColor: argbToRgb(styleFont?.colorArgb),
                      backgroundColor: argbToRgb(styleFill?.colorArgb),
                    },
                    layout: {
                      horizontalAlign: mapPLHorizontalAlignment(cf.style?.alignment?.horizontal),
                      verticalAlign: mapPLVerticalAlignment(cf.style?.alignment?.vertical),
                    },
                    behavior: {
                      required: cfg?.required || undefined,
                      readOnly: cfg?.readOnly || undefined,
                      visible: !(cfg?.hidden) || undefined,
                      enabled: undefined,
                    },
                    input: {
                      maxLength: (cfg?.maxLength ?? 0) > 0 ? cfg!.maxLength : undefined,
                      minLength: cfg?.minLength ?? undefined,
                    },
                  } as FieldConfig;

                  // Add keyboardText params so convertLegacyConfigToKtParams picks them up
                  if (runtimeField.dataType === "KeyboardText") {
                    const va = mapPLVerticalAlignment(cf.style?.alignment?.vertical);
                    const vaMap: Record<string, 0 | 1 | 2> = { top: 0, middle: 1, bottom: 2 };
                    const ha = mapPLHorizontalAlignment(cf.style?.alignment?.horizontal);
                    const haMap: Record<string, "Left" | "Center" | "Right"> = { left: "Left", center: "Center", right: "Right" };
                    const inputRestriction = cfg?.inputRestriction && cfg.inputRestriction !== "None" ? cfg.inputRestriction : "None";
                    runtimeField.config = {
                      ...runtimeField.config,
                      keyboardText: {
                        required: cfg?.required ?? false,
                        validateOnEditing: cfg?.validateOnEditing ?? false,
                        readOnly: cfg?.readOnly ?? false,
                        hidden: cfg?.hidden ?? false,
                        lines: cfg?.lines ?? 1,
                        inputRestriction: inputRestriction as any,
                        maxLength: cfg?.maxLength ?? 0,
                        align: (ha ? haMap[ha] : "Center") as any,
                        font: styleFont?.name || "Arial",
                        fontSize: (styleFont?.sizePt ?? 0) > 0 ? styleFont!.sizePt : 11,
                        defaultFontSize: (styleFont?.sizePt ?? 0) > 0 ? styleFont!.sizePt : 11,
                        weight: styleFont?.bold ? "Bold" : ("Normal" as any),
                        color: (() => {
                          const rgb = argbToRgb(styleFont?.colorArgb);
                          if (!rgb) return "0,0,0";
                          const c = rgb.replace("#", "");
                          if (c.length === 6) return `${parseInt(c.slice(0,2),16)},${parseInt(c.slice(2,4),16)},${parseInt(c.slice(4,6),16)}`;
                          return "0,0,0";
                        })(),
                        verticalAlignment: (va ? vaMap[va] : 1) as any,
                        placeholder: cfg?.placeholder ?? "",
                        defaultValue: cfg?.defaultValue ?? "",
                      }
                    } as FieldConfig;
                  }
                }
              }
            }

      // ═════════════════════════════════════════════════════════
      // PAPERLESS DEBUG STAGE 15 — After Re-upload Response
      // ═════════════════════════════════════════════════════════
      console.log("%c=========================================================", "color: #7c3aed; font-weight: bold");
      console.log("%cPAPERLESS DEBUG STAGE 15 — After Re-upload Response", "color: #7c3aed; font-weight: bold");
      console.log("%c=========================================================", "color: #7c3aed; font-weight: bold");
      {
        const s15Sheet = runtimeForm.sheets[0];
        const s15Field = s15Sheet?.fields[0];
        if (s15Field) {
          console.log("Field ID:", s15Field.id);
          console.log("Cell:", s15Field.cellReference);
          console.log("Font Name:", s15Field.font ?? "(not set)");
          console.log("Font Size:", s15Field.fontSize);
          console.log("Bold:", s15Field.bold);
          console.log("Font Color:", s15Field.fontColor ?? "(not set)");
          console.log("BG Color:", s15Field.backgroundColor ?? "(not set)");
          console.log("Alignment:", s15Field.alignment ?? "(not set)");
          console.log("Required:", s15Field.required);
          console.log("MaxLength:", s15Field.maxLength);
          console.log("DataType:", s15Field.dataType);
        }
      }
      console.log("%c=========================================================", "color: #7c3aed; font-weight: bold");
      console.log("");

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

  // ═════════════════════════════════════════════════════════
  // PHASE 21.5 — STAGE 2: React RuntimeForm State (on change)
  // ═════════════════════════════════════════════════════════
  useEffect(() => {
    if (runtimeForm) {
      console.log("%c=========================================================", "color: #0369a1; font-weight: bold");
      console.log("%cSTAGE 2 — React RuntimeForm State (after setRuntimeForm)", "color: #0369a1; font-weight: bold");
      console.log("%c=========================================================", "color: #0369a1; font-weight: bold");
      console.log("React state runtimeForm:");
      console.log("  sheets:", runtimeForm.sheets.length);
      runtimeForm.sheets.forEach((s, si) => {
        console.log(`  Sheet ${si}: '${s.name}' — Fields: ${s.fields.length}`);
        s.fields.forEach((f, fi) => {
          console.log(`    Field ${fi}: id='${f.id}' name='${f.name ?? "?"}' cell='${f.cellReference}'`);
        });
      });
      console.log("%c=========================================================", "color: #0369a1; font-weight: bold");
      console.log("");
    }
  }, [runtimeForm]);

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

  /**
   * Convert hex color #RRGGBB or #RGB to #AARRGGBB format.
   * Adds FF alpha prefix if not present.
   */
  function normalizeColorArgb(color: string | null | undefined): string | undefined {
    if (!color) return undefined;
    let c = color.replace("#", "");
    if (c.length === 3) c = c[0] + c[0] + c[1] + c[1] + c[2] + c[2];
    if (c.length === 6) return "#FF" + c.toUpperCase();
    if (c.length === 8) return "#" + c.toUpperCase();
    return undefined;
  }

  /**
   * Map frontend alignment string to backend HorizontalAlignment value.
   * "Center" → "center", "Left" → "left", "Right" → "right"
   */
  function mapHorizontalAlignment(align: string | null | undefined): string | undefined {
    if (!align) return undefined;
    const a = align.toLowerCase();
    if (a.includes("center")) return "center";
    if (a.includes("left")) return "left";
    if (a.includes("right")) return "right";
    return undefined;
  }

  /**
   * Map frontend vertical alignment to backend VerticalAlignment value.
   * "top" → "top", "middle" → "center", "bottom" → "bottom"
   */
  function mapVerticalAlignment(align: string | null | undefined): string | undefined {
    if (!align) return undefined;
    const a = align.toLowerCase();
    if (a === "top") return "top";
    if (a === "middle" || a === "center") return "center";
    if (a === "bottom") return "bottom";
    return undefined;
  }

  /**
   * Build WbDefStyle from RuntimeField + FieldConfig overrides.
   * FieldConfig overrides take priority over RuntimeField properties.
   */
  function buildFieldStyle(field: RuntimeField, config?: FieldConfig): WbDefStyle | undefined {
    const appearance = config?.appearance;
    const layout = config?.layout;

    // Read from RuntimeField properties (original upload values)
    let fontFamily = field.font ?? undefined;
    let fontSize = field.fontSize > 0 ? field.fontSize : undefined;
    let bold = field.bold || undefined;
    let fontColor = normalizeColorArgb(field.fontColor);
    let bgColor = normalizeColorArgb(field.backgroundColor);
    let hAlign = mapHorizontalAlignment(field.alignment);
    let vAlign = undefined as string | undefined;

    // Override with FieldConfig (user edits in browser)
    if (appearance) {
      if (appearance.fontFamily !== undefined) fontFamily = appearance.fontFamily;
      if (appearance.fontSize !== undefined) fontSize = appearance.fontSize;
      if (appearance.fontWeight !== undefined && appearance.fontWeight.toLowerCase() === "bold") bold = true;
      if (appearance.textColor !== undefined) fontColor = normalizeColorArgb(appearance.textColor);
      if (appearance.backgroundColor !== undefined) bgColor = normalizeColorArgb(appearance.backgroundColor);
    }
    if (layout) {
      if (layout.horizontalAlign !== undefined) hAlign = layout.horizontalAlign;
      if (layout.verticalAlign !== undefined) {
        // FieldConfig uses "middle", backend expects "center"
        vAlign = layout.verticalAlign === "middle" ? "center" : layout.verticalAlign;
      }
    }

    // Only emit style if there are actual changes from defaults
    const hasFontStyle = fontFamily !== undefined || fontSize !== undefined || bold !== undefined || fontColor !== undefined;
    const hasAlignment = hAlign !== undefined || vAlign !== undefined;
    const hasFill = bgColor !== undefined;

    if (!hasFontStyle && !hasAlignment && !hasFill) return undefined;

    const result: WbDefStyle = {};
    if (hasFontStyle) {
      result.font = {};
      if (fontFamily !== undefined) result.font.name = fontFamily;
      if (fontSize !== undefined) result.font.sizePt = fontSize;
      if (bold !== undefined) result.font.bold = bold;
      if (fontColor !== undefined) result.font.colorArgb = fontColor;
    }
    if (hasAlignment) {
      result.alignment = {};
      if (hAlign !== undefined) result.alignment.horizontal = hAlign;
      if (vAlign !== undefined) result.alignment.vertical = vAlign;
    }
    if (hasFill) {
      result.fill = {};
      result.fill.patternType = "solid";
      result.fill.colorArgb = bgColor;
    }
    return result;
  }

  // ── Convert RuntimeForm → WorkbookDefinition for save-edited flow (Phase 5.2) ──
  const runtimeFormToWorkbookDefinition = useCallback((
    form: RuntimeForm,
    values: Record<string, string | boolean | null>,
    sid: string,
    fieldConfigs?: Record<string, FieldConfig>,
  ): WbDef => {
    // ═════════════════════════════════════════════════════════
    // PHASE 21.5 — STAGE 6 & 7: runtimeFormToWorkbookDefinition
    // ═════════════════════════════════════════════════════════
    console.log("%c=========================================================", "color: #b45309; font-weight: bold");
    console.log("%cSTAGE 6 — runtimeFormToWorkbookDefinition INPUT", "color: #b45309; font-weight: bold");
    console.log("%c=========================================================", "color: #b45309; font-weight: bold");
    console.log("INPUT RuntimeForm:");
    console.log("  workbookName:", form.workbookName);
    console.log("  sheets:", form.sheets.length);
    form.sheets.forEach((s, si) => {
      console.log(`  Sheet ${si}: '${s.name}' — Fields: ${s.fields.length}`);
      s.fields.forEach((f, fi) => {
        const val = values[f.id];
        console.log(`    Field ${fi}: id='${f.id}' name='${f.name ?? "?"}' cell='${f.cellReference}' value='${val ?? "(empty)"}'`);
      });
    });

    // ═════════════════════════════════════════════════════════
    // BUG FIX (Phase 21.5): Include ALL fields in the payload,
    // not just fields with non-empty values. The filter was
    // removing EVERY field because values starts empty ({}).
    // Backend needs all fields to know which cells exist.
    // ═════════════════════════════════════════════════════════
    const result: WbDef = {
      info: { title: form.workbookName ?? "Untitled" },
      sourceFileName: sid,
      sessionId: sid,
      sheets: form.sheets.map((sheet, si) => ({
        name: sheet.name ?? `Page ${si + 1}`,
        index: si,
        fields: sheet.fields.map(f => {
          const fc = fieldConfigs?.[f.id];
          const beh = fc?.behavior ?? {};
          const inp = fc?.input ?? {};
          return {
            id: f.id,
            cell: {
              address: f.cellReference ?? "A1",
              rowIndex: parseInt(f.cellReference?.match(/\d+/)?.[0] ?? "1"),
            },
            name: f.name ?? "",
            // PHASE 21.6: Send integer enum value (0=Text, 1=Number, etc.)
            // instead of string. Backend uses System.Text.Json default which
            // expects integers for enums (no JsonStringEnumConverter).
            type: fieldTypeToBackendEnum(f.dataType ?? "KeyboardText"),
            value: String(values[f.id] ?? ""),
            // Phase 22: Browser style persistence — include user-edited styles
            style: buildFieldStyle(f, fieldConfigs?.[f.id]),
            // Full PaperLess config properties — sent so PaperLessConfigWriter
            // persists them in the hidden JSON for re-upload restoration.
            required: beh.required === true ? true : undefined,
            locked: beh.readOnly === true ? true : undefined,
            visible: beh.visible !== false ? undefined : false,
            maxLength: (inp.maxLength ?? 0) > 0 ? inp.maxLength : undefined,
            placeholder: fc?.input?.placeholder || undefined,
            defaultValue: fc?.input?.defaultValue || undefined,
            validateOnEditing: beh.validateOnEditing === true ? true : undefined,
          };
        }),
      })),
    };

    // ═════════════════════════════════════════════════════════
    // PHASE 21.5 — STAGE 7: OUTPUT PAYLOAD
    // ═════════════════════════════════════════════════════════
    console.log("%c=========================================================", "color: #b45309; font-weight: bold");
    console.log("%cSTAGE 7 — runtimeFormToWorkbookDefinition OUTPUT", "color: #b45309; font-weight: bold");
    console.log("%c=========================================================", "color: #b45309; font-weight: bold");
    console.log("OUTPUT WbDef:");
    console.log("  sheets:", result.sheets.length);
    result.sheets.forEach((s, si) => {
      console.log(`  Sheet ${si}: '${s.name}' — Fields: ${s.fields.length}`);
      s.fields.forEach((f, fi) => {
        console.log(`    Field ${fi}: id='${f.id}' name='${f.name}' cell='${f.cell.address}' value='${f.value}'`);
        if (f.style) {
          console.log(`    Style: font='${f.style.font?.name ?? "?"}/${f.style.font?.sizePt ?? "?"}' bold=${!!f.style.font?.bold} fill='${f.style.fill?.colorArgb ?? "none"}' hAlign='${f.style.alignment?.horizontal ?? "?"}'`);
        }
      });
    });
    console.log("%c=========================================================", "color: #b45309; font-weight: bold");
    console.log("");

    return result;
  }, []);

  // ── Save edited — POST /api/form/save-edited → WorkbookValueWriter (Phase 5.2) ──
  const handleSaveEdited = useCallback(async (fieldConfigs?: Record<string, FieldConfig>) => {
    if (!runtimeForm || !sessionId) {
      setExportError(sessionId ? "No form loaded." : "No session found. Please upload the workbook first.");
      return;
    }

    setExporting(true);
    setExportError(null);
    setExportSuccess(null);

    try {
      const wbDef = runtimeFormToWorkbookDefinition(runtimeForm, runtime.values, sessionId, fieldConfigs);

    // ═════════════════════════════════════════════════════════
    // PAPERLESS DEBUG STAGE 1 — Before Export Request
    // ═════════════════════════════════════════════════════════
    console.log("%c=========================================================", "color: #7c3aed; font-weight: bold");
    console.log("%cPAPERLESS DEBUG STAGE 1 — Before Export Request", "color: #7c3aed; font-weight: bold");
    console.log("%c=========================================================", "color: #7c3aed; font-weight: bold");
    {
      const s1Sheet = wbDef.sheets[0];
      const s1Field = s1Sheet?.fields[0];
      if (s1Field) {
        console.log("Field ID:", s1Field.id);
        console.log("Cell:", s1Field.cell.address);
        console.log("Font Name:", s1Field.style?.font?.name ?? "(not set)");
        console.log("Font Size:", s1Field.style?.font?.sizePt ?? "(not set)");
        console.log("Bold:", !!s1Field.style?.font?.bold);
        console.log("Italic:", !!s1Field.style?.font?.italic);
        console.log("Underline:", !!s1Field.style?.font?.underline);
        console.log("H-Align:", s1Field.style?.alignment?.horizontal ?? "(not set)");
      }
    }
    console.log("%c=========================================================", "color: #7c3aed; font-weight: bold");
    console.log("");

    // ═════════════════════════════════════════════════════════
    // PHASE 21.5 — STAGE 8: FINAL JSON PAYLOAD (before fetch)
    // ═════════════════════════════════════════════════════════
    // PHASE 22 — STAGE 22: Style payload diagnostic
    console.log("%c=========================================================", "color: #dc2626; font-weight: bold");
    console.log("%cSTAGE 8 — FINAL JSON PAYLOAD (before fetch)", "color: #dc2626; font-weight: bold");
    console.log("%c=========================================================", "color: #dc2626; font-weight: bold");
    console.log("FINAL FETCH PAYLOAD:");
    console.log("  sheets:", wbDef.sheets.length);
    wbDef.sheets.forEach((s, si) => {
      console.log(`  Sheet ${si}: '${s.name}' — Fields: ${s.fields.length}`);
      s.fields.forEach((f, fi) => {
        console.log(`    Field ${fi}: id='${f.id}' name='${f.name}' cell='${f.cell.address}' value='${f.value}'`);
        if (f.style) {
          console.log(`    Style: font='${f.style.font?.name ?? "?"}/${f.style.font?.sizePt ?? "?"}' bold=${!!f.style.font?.bold} fill='${f.style.fill?.colorArgb ?? "none"}' hAlign='${f.style.alignment?.horizontal ?? "?"}'`);
        }
      });
    });
    console.log("");
    console.log("JSON.stringify length:", JSON.stringify(wbDef).length, "bytes");
    console.log("fields[0] check: wbDef.sheets[0].fields.length =", wbDef.sheets[0]?.fields?.length ?? 0);
    console.log("%c=========================================================", "color: #dc2626; font-weight: bold");

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
  // Phase 22: Accept fieldConfigs from PaperlessDesigner for style persistence
  const handleExportExcel = useCallback(async (fieldConfigs?: Record<string, FieldConfig>) => {
    if (!runtimeForm) {
      setExportError("No form loaded.");
      return;
    }

    if (!sessionId) {
      setExportError("No session found. Please upload the workbook first.");
      return;
    }

    return handleSaveEdited(fieldConfigs);
  }, [runtimeForm, sessionId, handleSaveEdited]);

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
              onExportExcel={handleExportExcel as unknown as () => void}
              exporting={exporting}
              exportError={exportError}
              exportSuccess={exportSuccess}
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
