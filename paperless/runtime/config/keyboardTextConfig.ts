export const INPUT_RESTRICTIONS = ["None", "Numeric", "Alphabet", "Alphanumeric", "Katakana", "Hiragana"] as const;
export type InputRestriction = (typeof INPUT_RESTRICTIONS)[number];

export const ALIGN_OPTIONS = ["Left", "Center", "Right"] as const;
export type Align = (typeof ALIGN_OPTIONS)[number];

export const WEIGHT_OPTIONS = ["Normal", "Bold"] as const;
export type Weight = (typeof WEIGHT_OPTIONS)[number];

export type VerticalAlignment = 0 | 1 | 2;

export const VERTICAL_ALIGN_LABELS: Record<VerticalAlignment, string> = {
  0: "Top",
  1: "Center",
  2: "Bottom",
};

export interface KeyboardTextInputParameters {
  required: boolean;
  validateOnEditing: boolean;
  readOnly: boolean;
  hidden: boolean;
  lines: number;
  inputRestriction: InputRestriction;
  maxLength: number;
  align: Align;
  font: string;
  fontSize: number;
  defaultFontSize: number;
  weight: Weight;
  color: string;
  verticalAlignment: VerticalAlignment;
  placeholder: string;
  defaultValue: string;
}

export const DEFAULTS: KeyboardTextInputParameters = {
  required: false,
  validateOnEditing: false,
  readOnly: false,
  hidden: false,
  lines: 1,
  inputRestriction: "None",
  maxLength: 0,
  align: "Center",
  font: "Arial",
  fontSize: 11,
  defaultFontSize: 11,
  weight: "Normal",
  color: "0,0,0",
  verticalAlignment: 1,
  placeholder: "",
  defaultValue: "",
};

export function parseKeyboardTextInputParameters(
  params: Record<string, string>,
): KeyboardTextInputParameters {
  return {
    required: params["Required"] === "1",
    validateOnEditing: params["ValidateOnEditing"] === "1",
    readOnly: params["ReadOnly"] === "1",
    hidden: params["Hidden"] === "1",
    lines: Math.max(1, parseInt(params["Lines"], 10) || DEFAULTS.lines),
    inputRestriction: (INPUT_RESTRICTIONS as readonly string[]).includes(params["InputRestriction"])
      ? (params["InputRestriction"] as InputRestriction)
      : DEFAULTS.inputRestriction,
    maxLength: Math.max(0, parseInt(params["MaxLength"], 10) || 0),
    align: (ALIGN_OPTIONS as readonly string[]).includes(params["Align"])
      ? (params["Align"] as Align)
      : DEFAULTS.align,
    font: params["Font"] || DEFAULTS.font,
    fontSize: Math.max(1, parseInt(params["FontSize"], 10) || DEFAULTS.fontSize),
    defaultFontSize: Math.max(1, parseInt(params["DefaultFontSize"], 10) || DEFAULTS.defaultFontSize),
    weight: (WEIGHT_OPTIONS as readonly string[]).includes(params["Weight"])
      ? (params["Weight"] as Weight)
      : DEFAULTS.weight,
    color: params["Color"] || DEFAULTS.color,
    verticalAlignment: ([0, 1, 2] as number[]).includes(parseInt(params["VerticalAlignment"], 10))
      ? (parseInt(params["VerticalAlignment"], 10) as VerticalAlignment)
      : DEFAULTS.verticalAlignment,
    placeholder: params["Placeholder"] ?? DEFAULTS.placeholder,
    defaultValue: params["DefaultValue"] ?? DEFAULTS.defaultValue,
  };
}

export function keyboardTextToInputParametersString(params: KeyboardTextInputParameters): string {
  return [
    `Required=${params.required ? 1 : 0}`,
    `ValidateOnEditing=${params.validateOnEditing ? 1 : 0}`,
    `ReadOnly=${params.readOnly ? 1 : 0}`,
    `Hidden=${params.hidden ? 1 : 0}`,
    `Lines=${params.lines}`,
    `InputRestriction=${params.inputRestriction}`,
    `MaxLength=${params.maxLength}`,
    `Align=${params.align}`,
    `Font=${params.font}`,
    `FontSize=${params.fontSize}`,
    `DefaultFontSize=${params.defaultFontSize}`,
    `Weight=${params.weight}`,
    `Color=${params.color}`,
    `VerticalAlignment=${params.verticalAlignment}`,
    `Placeholder=${params.placeholder}`,
    `DefaultValue=${params.defaultValue}`,
  ].join(";") + ";";
}

export function rgbStringToHex(rgb: string): string {
  const parts = rgb.split(",").map((s) => parseInt(s.trim(), 10));
  if (parts.length !== 3 || parts.some(isNaN)) return "#1a1a1a";
  return "#" + parts.map((p) => Math.min(255, Math.max(0, p)).toString(16).padStart(2, "0")).join("");
}

export function hexToRgbString(hex: string): string {
  const clean = hex.replace("#", "");
  if (clean.length === 3) {
    return `${parseInt(clean[0] + clean[0], 16)},${parseInt(clean[1] + clean[1], 16)},${parseInt(clean[2] + clean[2], 16)}`;
  }
  if (clean.length === 6) {
    return `${parseInt(clean.slice(0, 2), 16)},${parseInt(clean.slice(2, 4), 16)},${parseInt(clean.slice(4, 6), 16)}`;
  }
  return "0,0,0";
}

const RESTRICTION_MAP: Record<string, string> = {
  None: "",
  Numeric: "numbers",
  Alphabet: "letters",
  Alphanumeric: "alphanumeric",
  Katakana: "",
  Hiragana: "",
};

const RESTRICTION_REVERSE: Record<string, InputRestriction> = {
  "": "None",
  numbers: "Numeric",
  letters: "Alphabet",
  alphanumeric: "Alphanumeric",
};

export function toCharacterRestriction(legacy: InputRestriction): string | undefined {
  return RESTRICTION_MAP[legacy] || undefined;
}

export function fromCharacterRestriction(current: string | undefined): InputRestriction {
  return RESTRICTION_REVERSE[current ?? ""] ?? "None";
}

const ALIGN_MAP: Record<string, "left" | "center" | "right"> = {
  Left: "left",
  Center: "center",
  Right: "right",
};

const ALIGN_REVERSE: Record<string, Align> = {
  left: "Left",
  center: "Center",
  right: "Right",
};

export function toHorizontalAlign(legacy: Align): "left" | "center" | "right" {
  return ALIGN_MAP[legacy] ?? "center";
}

export function fromHorizontalAlign(current: string): Align {
  return ALIGN_REVERSE[current] ?? "Center";
}

const WEIGHT_MAP: Record<string, string> = {
  Normal: "normal",
  Bold: "bold",
};

const WEIGHT_REVERSE: Record<string, Weight> = {
  normal: "Normal",
  bold: "Bold",
};

export function toFontWeight(legacy: Weight): string {
  return WEIGHT_MAP[legacy] ?? "normal";
}

export function fromFontWeight(current: string): Weight {
  return WEIGHT_REVERSE[current] ?? "Normal";
}

const VA_MAP: Record<number, "top" | "middle" | "bottom"> = {
  0: "top",
  1: "middle",
  2: "bottom",
};

const VA_REVERSE: Record<string, VerticalAlignment> = {
  top: 0,
  middle: 1,
  bottom: 2,
};

export function toVerticalAlign(legacy: VerticalAlignment): "top" | "middle" | "bottom" {
  return VA_MAP[legacy] ?? "middle";
}

export function fromVerticalAlign(current: string): VerticalAlignment {
  return VA_REVERSE[current] ?? 2;
}

export function toLines(multiline: boolean | undefined, heightPt?: number): number {
  if (multiline) return 2;
  if (heightPt && heightPt > 30) return 2;
  return 1;
}

export function fromLines(lines: number): boolean {
  return lines > 1;
}
