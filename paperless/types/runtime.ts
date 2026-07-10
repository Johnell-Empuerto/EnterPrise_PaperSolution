/**
 * TypeScript type definitions for the FormLess Runtime Engine (Phase 11I/11J).
 * Mirrors the C# RuntimeForm/RuntimeField/RuntimeSheet models from the backend.
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
}

export interface RuntimeField {
  id: string;
  cellReference: string;
  row: number;
  column: number;
  leftPx: number;
  topPx: number;
  widthPx: number;
  heightPx: number;
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
