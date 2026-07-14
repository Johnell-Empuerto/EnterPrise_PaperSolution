"use client";

/**
 * BackgroundLayer — renders the PDF-generated PNG at native pixel dimensions.
 *
 * Behavior:
 * - Positioned at (0,0) within the PageSurface
 * - Sizes to fill the layer (which is 100% of PageSurface)
 * - maxWidth: "none" — overrides Tailwind preflight img { max-width: 100% }
 * - display: block — removes inline gap below image
 * - No object-fit, no transforms, no scaling — Stretch=None equivalent
 *
 * Legacy equivalent: WPF Image with Stretch=None, Width/Height set to bitmap size.
 */
export interface BackgroundLayerProps {
  /** Background image URL (PNG rendered from Excel / COM backend) */
  src: string;
  /** Alt text for accessibility */
  alt: string;
  /** Native width of the PNG in pixels */
  widthPx: number;
  /** Native height of the PNG in pixels */
  heightPx: number;
}

export function BackgroundLayer({ src, alt, widthPx, heightPx }: BackgroundLayerProps) {
  return (
    <div
      data-background-layer
      style={{
        position: "absolute",
        left: 0,
        top: 0,
        width: "100%",
        height: "100%",
      }}
    >
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img
        src={src}
        alt={alt}
        data-form-background
        style={{
          display: "block",
          width: widthPx,
          height: heightPx,
          maxWidth: "none", // CRITICAL: override Tailwind preflight max-width:100%
        }}
        draggable={false}
      />
    </div>
  );
}
