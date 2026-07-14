import type { LegacyRuntimeField } from "./LegacyCanvasRuntime";

export interface HitResult {
  field: LegacyRuntimeField;
  index: number;
}

export function hitTest(
  fields: LegacyRuntimeField[],
  pageX: number,
  pageY: number
): HitResult | null {
  for (let i = fields.length - 1; i >= 0; i--) {
    const f = fields[i];
    if (
      pageX >= f.leftPx &&
      pageX <= f.leftPx + f.widthPx &&
      pageY >= f.topPx &&
      pageY <= f.topPx + f.heightPx
    ) {
      return { field: f, index: i };
    }
  }
  return null;
}
