import type { RuntimeField, FieldConfig } from "@/types/runtime";
import type { KeyboardTextInputParameters, InputRestriction } from "./keyboardTextConfig";
import { DEFAULTS } from "./keyboardTextConfig";

export function isLegacyTextField(field: RuntimeField): boolean {
  const dt = field.dataType.toLowerCase();
  return dt === "text" || dt === "textbox";
}

export function convertLegacyConfigToKtParams(config: Record<string, any> | undefined): KeyboardTextInputParameters {
  if (!config) return { ...DEFAULTS };

  const input = (config.input ?? {}) as Record<string, any>;
  const behavior = (config.behavior ?? {}) as Record<string, any>;
  const appearance = (config.appearance ?? {}) as Record<string, any>;
  const layout = (config.layout ?? {}) as Record<string, any>;

  const rawRestriction: string | undefined = input.characterRestriction;
  const inputRestriction: InputRestriction = rawRestriction
    ? mapCharacterRestriction(rawRestriction)
    : "None";

  const verticalAlignment = mapVerticalAlignment(layout.verticalAlign);
  const align = mapHorizontalAlign(layout.horizontalAlign);
  const weight = mapFontWeight(appearance.fontWeight);
  const color = hexToRgbString(appearance.textColor ?? "#000000");

  return {
    required: behavior.required ?? false,
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
    placeholder: (config.placeholder as string) ?? "",
    defaultValue: (config.defaultValue as string) ?? "",
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
  return va ? (map[va] ?? 2) : 2;
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
