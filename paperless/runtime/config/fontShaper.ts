"use client";

export interface ShapeOptions {
  fontFamily: string;
  fontWeight: string;
  originalFontSizePt: number;
  minFontSizePt: number;
  lineHeightRatio: number;
  availableWidthPx: number;
  availableHeightPx: number;
}

export interface ShapeResult {
  fontSize: number;
  lines: string[];
  contentHeightPx: number;
}

let _canvas: HTMLCanvasElement | null = null;
let _ctx: CanvasRenderingContext2D | null = null;

function ctx(): CanvasRenderingContext2D {
  if (!_canvas) {
    _canvas = document.createElement("canvas");
    _ctx = _canvas.getContext("2d");
  }
  if (!_ctx) throw new Error("Canvas 2D context unavailable");
  return _ctx;
}

export function shapeText(text: string, opts: ShapeOptions): ShapeResult {
  const PT_PX = 96 / 72;

  function measure(sizePt: number): { lines: string[]; contentHeightPx: number } {
    const lines = wrap(text, opts.fontFamily, opts.fontWeight, sizePt, opts.availableWidthPx);
    const lh = sizePt * PT_PX * opts.lineHeightRatio;
    return { lines, contentHeightPx: lines.length * lh };
  }

  if (!text || opts.availableWidthPx <= 0) {
    const h = opts.originalFontSizePt * PT_PX * opts.lineHeightRatio;
    return { fontSize: opts.originalFontSizePt, lines: [""], contentHeightPx: h };
  }

  const orig = measure(opts.originalFontSizePt);
  if (orig.contentHeightPx <= opts.availableHeightPx) {
    return { fontSize: opts.originalFontSizePt, ...orig };
  }

  const STEP = 0.5;
  let cur = opts.originalFontSizePt - STEP;
  while (cur >= opts.minFontSizePt) {
    const r = measure(cur);
    if (r.contentHeightPx <= opts.availableHeightPx) {
      return { fontSize: cur, ...r };
    }
    cur -= STEP;
  }

  const fallback = measure(opts.minFontSizePt);
  return { fontSize: opts.minFontSizePt, ...fallback };
}

function wrap(
  text: string,
  fontFamily: string,
  fontWeight: string,
  fontSizePt: number,
  maxW: number,
): string[] {
  const c = ctx();
  c.font = `${fontWeight} ${fontSizePt}pt ${fontFamily}`;

  const lines: string[] = [];
  for (const para of text.split("\n")) {
    if (para === "") {
      lines.push("");
      continue;
    }
    const words = para.split(" ");
    let cur = "";
    for (const w of words) {
      if (!w) continue;

      // If a single word is wider than maxW, break it character-by-character
      if (c.measureText(w).width > maxW) {
        if (cur) {
          lines.push(cur);
          cur = "";
        }
        let charLine = "";
        for (const ch of w) {
          const cand = charLine + ch;
          if (c.measureText(cand).width > maxW && charLine) {
            lines.push(charLine);
            charLine = ch;
          } else {
            charLine = cand;
          }
        }
        if (charLine) cur = charLine;
        continue;
      }

      const cand = cur ? cur + " " + w : w;
      if (c.measureText(cand).width > maxW && cur) {
        lines.push(cur);
        cur = w;
      } else {
        cur = cand;
      }
    }
    if (cur) lines.push(cur);
  }
  return lines;
}

export const LINE_HEIGHT_RATIO = 1.4;
export const MIN_FONT_SIZE_PT = 6;
export const MEASURE_TOLERANCE_PX = 2;
