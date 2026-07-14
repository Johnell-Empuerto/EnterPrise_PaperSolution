"use client";

import { useMemo } from "react";
import type { TemplateModel } from "@/types/template";
import { ExcelGrid } from "./ExcelGrid";
import { originFromMargins, pt } from "./helpers";
import styles from "./ExcelRenderer.module.css";

export interface ExcelPageProps {
  template: TemplateModel;
}

export function ExcelPage({ template }: ExcelPageProps) {
  const { pageSetup } = template;

  const origin = useMemo(() => originFromMargins(template), [template]);

  const pageStyle: React.CSSProperties = {
    width: pt(pageSetup.paperWidthPt),
    height: pt(pageSetup.paperHeightPt),
    padding: `${pt(pageSetup.marginTopIn * 72)} ${pt(pageSetup.marginRightIn * 72)} ${pt(pageSetup.marginBottomIn * 72)} ${pt(pageSetup.marginLeftIn * 72)}`,
    backgroundColor: "#ffffff",
  };

  const gridWrapperStyle: React.CSSProperties = {
    position: "relative",
    marginLeft: pageSetup.centerHorizontally ? "auto" : undefined,
    marginRight: pageSetup.centerHorizontally ? "auto" : undefined,
  };

  return (
    <div className={styles.page} style={pageStyle} data-excel-page>
      <div style={gridWrapperStyle}>
        <ExcelGrid template={template} />
      </div>
    </div>
  );
}
