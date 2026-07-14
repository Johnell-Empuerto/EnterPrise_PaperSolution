"use client";

import { useMemo } from "react";
import type { OverlayCollection } from "@/types/overlay";

export interface RuntimeInspectorProps {
  overlayCollection: OverlayCollection;
  values: Record<string, string | boolean | null>;
  dirty: Record<string, boolean>;
  fieldCount: number;
  onCopyJson: () => void;
  onReset: () => void;
}

const INTERACTIVE_TYPES = new Set(["textbox", "checkbox", "date", "number", "signature"]);

export function RuntimeInspector({
  overlayCollection,
  values,
  dirty,
  fieldCount,
  onCopyJson,
  onReset,
}: RuntimeInspectorProps) {
  const interactive = useMemo(
    () => overlayCollection.overlays.filter((o) => INTERACTIVE_TYPES.has(o.type)),
    [overlayCollection]
  );

  return (
    <div className="text-xs">
      <div className="flex items-center justify-between mb-2 pb-2 border-b border-gray-200">
        <span className="font-medium text-gray-600">
          {interactive.length} fields · {fieldCount} filled
        </span>
        <div className="flex gap-1.5">
          <button
            className="px-2 py-0.5 text-[10px] rounded bg-gray-100 text-gray-600 hover:bg-gray-200"
            onClick={onCopyJson}
            title="Copy runtime values as JSON"
          >
            Copy Values
          </button>
          <button
            className="px-2 py-0.5 text-[10px] rounded bg-red-50 text-red-500 hover:bg-red-100"
            onClick={onReset}
            title="Reset all values"
          >
            Reset
          </button>
        </div>
      </div>

      {/* Field list */}
      <div className="max-h-60 overflow-y-auto space-y-1">
        {interactive.map((overlay) => {
          const val = values[overlay.id];
          const isDirty = dirty[overlay.id];
          const displayValue =
            val === null || val === undefined
              ? "" : typeof val === "boolean"
              ? val ? "✓" : "✗" : (val as string);

          return (
            <div
              key={overlay.id}
              className={`flex items-center gap-2 py-0.5 px-1 rounded ${
                isDirty ? "bg-amber-50" : "bg-gray-50/50"
              }`}
            >
              {/* Type badge */}
              <span
                className={`text-[9px] px-1 rounded font-medium ${
                  typeBadgeColor(overlay.type)
                }`}
                style={{ minWidth: 28, textAlign: "center" }}
              >
                {overlay.type.slice(0, 4).toUpperCase()}
              </span>

              {/* ID + Cell */}
              <span className="font-mono text-[10px] text-gray-500 min-w-[60px]">
                {overlay.cell}
              </span>

              {/* Coordinates */}
              <span className="font-mono text-[9px] text-gray-400 hidden sm:inline">
                {overlay.leftPt.toFixed(0)},{overlay.topPt.toFixed(0)}
              </span>

              {/* Value */}
              <span className="font-mono text-[10px] text-gray-700 truncate flex-1">
                {displayValue || <span className="text-gray-300 italic">empty</span>}
              </span>

              {/* Dirty indicator */}
              {isDirty && (
                <span className="text-[8px] text-amber-500 font-bold">●</span>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function typeBadgeColor(type: string): string {
  const colors: Record<string, string> = {
    textbox: "bg-blue-100 text-blue-700",
    signature: "bg-purple-100 text-purple-700",
    checkbox: "bg-green-100 text-green-700",
    date: "bg-amber-100 text-amber-700",
    number: "bg-red-100 text-red-700",
  };
  return colors[type] ?? "bg-gray-100 text-gray-700";
}
