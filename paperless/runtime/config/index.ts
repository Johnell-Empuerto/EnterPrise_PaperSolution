export {
  parseKeyboardTextInputParameters,
  keyboardTextToInputParametersString,
  rgbStringToHex,
  hexToRgbString,
  toCharacterRestriction,
  fromCharacterRestriction,
  toHorizontalAlign,
  fromHorizontalAlign,
  toVerticalAlign,
  fromVerticalAlign,
  toFontWeight,
  fromFontWeight,
  toLines,
  fromLines,
  DEFAULTS,
} from "./keyboardTextConfig";
export type {
  KeyboardTextInputParameters,
  InputRestriction,
  Align,
  Weight,
  VerticalAlignment,
} from "./keyboardTextConfig";
export {
  migrateFieldToKeyboardText,
  isLegacyTextField,
  convertLegacyConfigToKtParams,
} from "./migration";
