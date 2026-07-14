export interface PagePoint {
  x: number;
  y: number;
}

export function canvasPoint(
  canvas: HTMLCanvasElement,
  clientX: number,
  clientY: number
): PagePoint {
  const rect = canvas.getBoundingClientRect();
  return {
    x: clientX - rect.left,
    y: clientY - rect.top,
  };
}
