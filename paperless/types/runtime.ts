/**
 * TypeScript type definitions for the FormLess Runtime Engine (Phase 11I/11J).
 * Mirrors the C# RuntimeForm/RuntimeField/RuntimeSheet models from the backend.
 *
 * Phase 14: Added ratio-based coordinates (leftRatio, topRatio, widthRatio, heightRatio)
 * for legacy ConMas compatibility. Frontend falls back to ratio-based rendering when
 * pixel coordinates are missing or zero.
 *
 * Consumed by GET /api/runtime/{templateId}
 */

export interface RuntimeForm {
  workbookName: string;
  title: string;
  sheets: RuntimeSheet[];
  pageWidth: number;
  pageHeight: number;
  scale: number;
  dpi: number;
  version: string;
}

export interface RuntimeSheet {
  name: string;
  index: number;
  fields: RuntimeField[];
  images: RuntimeImage[];
  shapes: RuntimeShape[];
  printArea: RuntimePrintArea | null;
  pageWidthPx: number;
  pageHeightPx: number;
  /**
   * Background image URL (relative, e.g., "/forms/bg_xxx.png").
   * This is the PNG rendered from Excel — the exact same image used during
   * coordinate computation. The frontend MUST use this to ensure the
   * background image matches the field coordinates.
   * Falls back to constructing from template ID if not provided.
   */
  backgroundImage?: string | null;
}

export interface FieldConfig {
  appearance: {
    fontFamily?: string;
    fontSize?: number;
    fontWeight?: string;
    textColor?: string;
    backgroundColor?: string;
    border?: string;
    borderRadius?: string;
    textAlign?: string;
  };
  behavior: {
    readOnly?: boolean;
    required?: boolean;
    visible?: boolean;
    enabled?: boolean;
    multiline?: boolean;
  };
  input: {
    keyboardType?: "text" | "number" | "decimal" | "email" | "phone" | "password" | "url";
    characterRestriction?: string;
    maxLength?: number;
    minLength?: number;
  };
  layout: {
    horizontalAlign?: "left" | "center" | "right";
    verticalAlign?: "top" | "middle" | "bottom";
  };
}

export interface RuntimeField {
  id: string;
  /** User-visible field name (from comment or default). Never shown as `id`. */
  name?: string;
  cellReference: string;
  row: number;
  column: number;
  leftPx: number;
  topPx: number;
  widthPx: number;
  heightPx: number;
  /** Left edge as ratio of page width (0-1) — for legacy ConMas compatibility. */
  leftRatio: number;
  /** Top edge as ratio of page height (0-1) — for legacy ConMas compatibility. */
  topRatio: number;
  /** Width as ratio of page width (0-1) — for legacy ConMas compatibility. */
  widthRatio: number;
  /** Height as ratio of page height (0-1) — for legacy ConMas compatibility. */
  heightRatio: number;
  mergeRange: string | null;
  isMerged: boolean;
  dataType: "text" | "number" | "date" | "checkbox" | "signature" | "dropdown" | "calculated";
  readOnly: boolean;
  required: boolean;
  alignment: string | null;
  font: string | null;
  fontSize: number;
  bold: boolean;
  fontColor: string | null;
  backgroundColor: string | null;
  border: string | null;
  placeholder: string | null;
  defaultValue: string | null;
  maxLength: number;
  tabIndex: number;
  validationPattern?: string | null;
  validationMessage?: string | null;
  /** Live-editable configuration overrides — source of truth during editing. */
  config?: FieldConfig;
}

export interface RuntimeImage {
  name: string;
  leftPx: number;
  topPx: number;
  widthPx: number;
  heightPx: number;
  contentType: string;
}

export interface RuntimeShape {
  name: string;
  shapeType: string;
  leftPx: number;
  topPx: number;
  widthPx: number;
  heightPx: number;
}

export interface RuntimePrintArea {
  address: string;
  leftPx: number;
  topPx: number;
  widthPx: number;
  heightPx: number;
}

export interface FieldValues {
  [fieldId: string]: string | boolean | null;
}

export interface FieldErrors {
  [fieldId: string]: string | null;
}

export interface DirtyState {
  [fieldId: string]: boolean;
}

/**
 * Phase 29 — Interactive Form Runtime
 * Runtime value model for storing and exporting form field values.
 */
export interface RuntimeValueModel {
  overlayId: string;
  value: string | boolean | null;
}

/**
 * The complete runtime state for a form session.
 */
export interface RuntimeSessionState {
  /** Values keyed by overlay id */
  values: FieldValues;
  /** Track which fields have been modified */
  dirty: DirtyState;
  /** ISO timestamp of last update */
  lastUpdated: string | null;
}
