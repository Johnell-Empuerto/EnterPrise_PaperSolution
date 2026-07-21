import type { RuntimeField, FieldConfig } from "@/types/runtime";
import type { KeyboardTextInputParameters, InputRestriction } from "./keyboardTextConfig";
import { DEFAULTS } from "./keyboardTextConfig";

export function isLegacyTextField(field: RuntimeField): boolean {
  const dt = field.dataType.toLowerCase();
  return dt === "text" || dt === "textbox";
}

export function convertLegacyConfigToKtParams(config: Record<string, any> | undefined): KeyboardTextInputParameters {
  if (!config) {
    // [PaperLess Upload Debug] PART 8 — convertLegacyConfigToKtParams (NO config)
    console.groupCollapsed(`[PaperLess Upload Debug] convertLegacyConfigToKtParams — NO CONFIG`);
    console.log("Returning DEFAULTS:", { ...DEFAULTS });
    console.groupEnd();
    return { ...DEFAULTS };
  }

  const input = (config.input ?? {}) as Record<string, any>;
  const behavior = (config.behavior ?? {}) as Record<string, any>;
  const appearance = (config.appearance ?? {}) as Record<string, any>;
  const layout = (config.layout ?? {}) as Record<string, any>;

  // Read from canonical config.input.inputRestriction (set by restoration),
  // then config.keyboardText.inputRestriction (legacy restoration path),
  // then config.input.characterRestriction (legacy/non-canonical path).
  const rawRestriction: string | undefined = input.inputRestriction
    ?? (config as any).keyboardText?.inputRestriction
    ?? input.characterRestriction;
  const inputRestriction: InputRestriction = rawRestriction
    ? mapCharacterRestriction(rawRestriction)
    : "None";

  const verticalAlignment = mapVerticalAlignment(layout.verticalAlign);
  const align = mapHorizontalAlign(layout.horizontalAlign);
  const weight = mapFontWeight(appearance.fontWeight);
  const color = hexToRgbString(appearance.textColor ?? "#000000");

  // [PaperLess Upload Debug] PART 8 — convertLegacyConfigToKtParams
  console.groupCollapsed(`[PaperLess Upload Debug] convertLegacyConfigToKtParams`);
  console.log("Raw config:", config);
  console.log("  appearance:", appearance);
  console.log("  layout:", layout);
  console.log("  behavior:", behavior);
  console.log("  input:", input);
  console.log("  characterRestriction:", input.characterRestriction);
  console.log("  inputRestriction (mapped):", inputRestriction);
  console.log("  layout.verticalAlign:", layout.verticalAlign);
  console.log("  layout.horizontalAlign:", layout.horizontalAlign);
  console.log("  verticalAlignment (mapped):", verticalAlignment);
  console.log("  horizontalAlign (mapped):", align);
  console.log("  appearance.fontWeight:", appearance.fontWeight);
  console.log("  weight (mapped):", weight);
  console.log("  appearance.textColor:", appearance.textColor);
  console.log("  color (mapped):", color);
  console.log("  config.keyboardText?.placeholder:", config.keyboardText?.placeholder);
  console.log("  input.placeholder:", input.placeholder);
  console.log("  config.placeholder:", config.placeholder);
  console.log("  config.keyboardText?.defaultValue:", config.keyboardText?.defaultValue);
  console.log("  input.defaultValue:", input.defaultValue);
  console.log("  config.defaultValue:", config.defaultValue);
  console.log("Result:", {
    align, inputRestriction, verticalAlignment, weight, color,
    font: appearance.fontFamily?.split(",")[0]?.trim() ?? DEFAULTS.font,
    fontSize: appearance.fontSize ?? DEFAULTS.fontSize,
    placeholder: (config.keyboardText?.placeholder as string) ?? (input.placeholder as string) ?? (config.placeholder as string) ?? "",
    defaultValue: (config.keyboardText?.defaultValue as string) ?? (input.defaultValue as string) ?? (config.defaultValue as string) ?? "",
  });
  console.groupEnd();

  return {
    required: behavior.required ?? false,
    validateOnEditing: false,
    readOnly: behavior.readOnly ?? false,
    hidden: behavior.hidden === true || behavior.hidden === "1",
    lines: behavior.multiline ? 2 : 1,
    inputRestriction,
    maxLength: input.maxLength ?? 0,
    align,
    font: appearance.fontFamily?.split(",")[0]?.trim() ?? DEFAULTS.font,
    fontSize: appearance.fontSize ?? DEFAULTS.fontSize,
    defaultFontSize: appearance.fontSize ?? DEFAULTS.defaultFontSize,
    weight,
    color,
    verticalAlignment,
    // Read from config.keyboardText (set by PaperLessConfig restoration),
    // then config.input (set by PaperLessConfig restoration), then config top-level (legacy).
    // This ensures placeholder/defaultValue survive re-upload round trips.
    placeholder: (config.keyboardText?.placeholder as string) ?? (input.placeholder as string) ?? (config.placeholder as string) ?? "",
    defaultValue: (config.keyboardText?.defaultValue as string) ?? (input.defaultValue as string) ?? (config.defaultValue as string) ?? "",
  };
}

export function migrateFieldToKeyboardText(field: RuntimeField): RuntimeField {
  if (field.dataType === "KeyboardText") return field;
  if (!isLegacyTextField(field)) return field;

  const ktParams = convertLegacyConfigToKtParams(field.config);

  return {
    ...field,
    dataType: "KeyboardText",
    config: {
      ...field.config,
      keyboardText: ktParams,
    } as FieldConfig,
  };
}

export function migrateFieldToForward(
  field: RuntimeField,
): RuntimeField {
  if (field.dataType !== "KeyboardText") return field;
  return field;
}

function mapCharacterRestriction(legacy: string): InputRestriction {
  const map: Record<string, InputRestriction> = {
    "": "None",
    none: "None",
    letters: "Alphabet",
    numbers: "Numeric",
    numeric: "Numeric",
    alphabet: "Alphabet",
    alphanumeric: "Alphanumeric",
    uppercase: "Alphabet",
    lowercase: "Alphabet",
  };
  return map[legacy.toLowerCase()] ?? "None";
}

function mapVerticalAlignment(va: string | undefined): 0 | 1 | 2 {
  const map: Record<string, 0 | 1 | 2> = {
    top: 0,
    middle: 1,
    bottom: 2,
  };
  return va ? (map[va] ?? 1) : 1;
}

function mapHorizontalAlign(ha: string | undefined): "Left" | "Center" | "Right" {
  const map: Record<string, "Left" | "Center" | "Right"> = {
    left: "Left",
    center: "Center",
    right: "Right",
  };
  return ha ? (map[ha] ?? "Center") : "Center";
}

function mapFontWeight(fw: string | undefined): "Normal" | "Bold" {
  if (!fw) return "Normal";
  const val = fw.toLowerCase();
  if (val === "bold" || val === "700" || val === "800" || val === "900") return "Bold";
  return "Normal";
}

function hexToRgbString(hex: string): string {
  const clean = hex.replace("#", "");
  if (clean.length === 3) {
    return `${parseInt(clean[0] + clean[0], 16)},${parseInt(clean[1] + clean[1], 16)},${parseInt(clean[2] + clean[2], 16)}`;
  }
  if (clean.length === 6) {
    return `${parseInt(clean.slice(0, 2), 16)},${parseInt(clean.slice(2, 4), 16)},${parseInt(clean.slice(4, 6), 16)}`;
  }
  return "0,0,0";
}
