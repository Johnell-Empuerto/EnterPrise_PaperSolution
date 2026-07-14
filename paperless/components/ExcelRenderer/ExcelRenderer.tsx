"use client";

import type { TemplateModel } from "@/types/template";
import { ExcelPage } from "./ExcelPage";

export interface ExcelRendererProps {
  template: TemplateModel;
}

export function ExcelRenderer({ template }: ExcelRendererProps) {
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
      <ExcelPage template={template} />
    </div>
  );
}
