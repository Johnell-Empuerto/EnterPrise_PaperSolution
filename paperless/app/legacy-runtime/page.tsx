import LegacyCanvasRuntime from "@/components/LegacyCanvasRuntime/LegacyCanvasRuntime";

export const metadata = {
  title: "Legacy Canvas Runtime",
};

export default async function LegacyRuntimePage({
  searchParams,
}: {
  searchParams: Promise<{ templateId?: string }>;
}) {
  const params = await searchParams;
  const templateId = params?.templateId ?? "547";
  return <LegacyCanvasRuntime templateId={templateId} />;
}
