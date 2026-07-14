import type { TemplateModel } from "@/types/template";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:5090";

export async function fetchTemplate(id: number): Promise<TemplateModel> {
  const res = await fetch(`${API_BASE_URL}/api/template/${id}`);
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`HTTP ${res.status}: ${body}`);
  }
  return res.json();
}

export const TEMPLATE_IDS = [546, 547, 548] as const;

export type TemplateId = (typeof TEMPLATE_IDS)[number];
