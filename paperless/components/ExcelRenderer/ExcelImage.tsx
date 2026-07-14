"use client";

import type { TemplateImage } from "@/types/template";
import styles from "./ExcelRenderer.module.css";

export interface ExcelImageProps {
  image: TemplateImage;
  leftPt: number;
  topPt: number;
}

export function ExcelImage({ image, leftPt, topPt }: ExcelImageProps) {
  const src = image.imageData ?? image.imageUrl;
  if (!src) return null;

  const inlineStyle: React.CSSProperties = {
    position: "absolute",
    left: `${leftPt}pt`,
    top: `${topPt}pt`,
    width: `${image.widthPt}pt`,
    height: `${image.heightPt}pt`,
  };

  return (
    <div className={styles.image} style={inlineStyle}>
      <img
        src={src}
        alt={image.description ?? "Embedded image"}
        draggable={false}
        style={{
          width: "100%",
          height: "100%",
          objectFit: "contain",
          objectPosition: "top left",
          pointerEvents: "none",
          userSelect: "none",
        }}
      />
    </div>
  );
}
