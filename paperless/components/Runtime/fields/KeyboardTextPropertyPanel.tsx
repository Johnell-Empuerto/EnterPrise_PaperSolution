"use client";

import type { KeyboardTextInputParameters, InputRestriction, Align, Weight, VerticalAlignment } from "@/runtime/config/keyboardTextConfig";
import {
  INPUT_RESTRICTIONS,
  ALIGN_OPTIONS,
  WEIGHT_OPTIONS,
  VERTICAL_ALIGN_LABELS,
  rgbStringToHex,
  hexToRgbString,
} from "@/runtime/config/keyboardTextConfig";

export interface KeyboardTextPropertyPanelProps {
  params: KeyboardTextInputParameters;
  onChange: (params: KeyboardTextInputParameters) => void;
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <div className="text-[10px] text-slate-400 uppercase tracking-wide mb-0.5 font-medium">{label}</div>
      {children}
    </div>
  );
}

function Section({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="border border-slate-100 rounded-md overflow-hidden">
      <div className="px-2.5 py-1.5 text-xs font-semibold text-slate-700 bg-slate-50 border-b border-slate-100">
        {label}
      </div>
      <div className="px-2.5 py-2 space-y-2">
        {children}
      </div>
    </div>
  );
}

function Toggle({
  label,
  checked,
  onChange,
}: {
  label: string;
  checked: boolean;
  onChange: (v: boolean) => void;
}) {
  return (
    <label className="flex items-center justify-between cursor-pointer">
      <span className="text-xs text-slate-600">{label}</span>
      <button
        onClick={() => onChange(!checked)}
        className={`relative w-8 h-4 rounded-full transition-colors ${
          checked ? "bg-emerald-500" : "bg-slate-300"
        }`}
      >
        <span
          className={`absolute top-0.5 left-0.5 w-3 h-3 bg-white rounded-full transition-transform ${
            checked ? "translate-x-4" : ""
          }`}
        />
      </button>
    </label>
  );
}

function Select({
  value,
  onChange,
  options,
}: {
  value: string;
  onChange: (v: string) => void;
  options: { value: string; label: string }[];
}) {
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      className="w-full text-xs text-slate-700 px-2 py-1 border border-slate-200 rounded-md focus:outline-none focus:border-indigo-400 bg-white"
    >
      {options.map((opt) => (
        <option key={opt.value} value={opt.value}>
          {opt.label}
        </option>
      ))}
    </select>
  );
}

function NumberInput({
  value,
  min,
  placeholder,
  onChange,
}: {
  value: number;
  min: number;
  placeholder?: string;
  onChange: (v: number) => void;
}) {
  return (
    <input
      type="number"
      min={min}
      value={value}
      onChange={(e) => onChange(Math.max(min, Number(e.target.value) || 0))}
      placeholder={placeholder}
      className="w-full text-xs text-slate-700 px-2 py-1 border border-slate-200 rounded-md focus:outline-none focus:border-indigo-400 bg-white"
    />
  );
}

function TextInput({
  value,
  placeholder,
  onChange,
}: {
  value: string;
  placeholder?: string;
  onChange: (v: string) => void;
}) {
  return (
    <input
      type="text"
      value={value}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
      className="w-full text-xs text-slate-700 px-2 py-1 border border-slate-200 rounded-md focus:outline-none focus:border-indigo-400 bg-white"
    />
  );
}

function SegmentedControl<T extends string | number>({
  value,
  onChange,
  options,
}: {
  value: T;
  onChange: (v: T) => void;
  options: { value: T; label: string }[];
}) {
  return (
    <div className="flex gap-1">
      {options.map((opt) => (
        <button
          key={String(opt.value)}
          onClick={() => onChange(opt.value)}
          className={`flex-1 text-[10px] px-1 py-1 rounded border transition-colors ${
            value === opt.value
              ? "bg-indigo-100 border-indigo-400 text-indigo-700 font-medium"
              : "border-slate-200 text-slate-500 hover:bg-slate-50"
          }`}
        >
          {opt.label}
        </button>
      ))}
    </div>
  );
}

export function KeyboardTextPropertyPanel({ params, onChange }: KeyboardTextPropertyPanelProps) {
  const update = <K extends keyof KeyboardTextInputParameters>(
    key: K,
    value: KeyboardTextInputParameters[K],
  ) => {
    onChange({ ...params, [key]: value });
  };

  const fontOptions = [
    { value: "Arial", label: "Arial" },
    { value: "Calibri", label: "Calibri" },
    { value: "Times New Roman", label: "Times New Roman" },
    { value: "Courier New", label: "Courier New" },
    { value: "Georgia", label: "Georgia" },
    { value: "Verdana", label: "Verdana" },
  ];

  const fontSizeOptions = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72].map(
    (s) => ({ value: String(s), label: `${s}pt` }),
  );

  const horizontals: { value: Align; label: string }[] = ALIGN_OPTIONS.map((a) => ({
    value: a,
    label: a,
  }));

  const verticals: { value: VerticalAlignment; label: string }[] = ([0, 1, 2] as VerticalAlignment[]).map(
    (va) => ({ value: va, label: VERTICAL_ALIGN_LABELS[va] }),
  );

  const restrictionOptions = INPUT_RESTRICTIONS.map((r) => ({
    value: r,
    label: r,
  }));

  return (
    <div className="space-y-2">
      {/* ── Font ── */}
      <Section label="Font">
        <Field label="Font">
          <Select
            value={params.font}
            onChange={(v) => update("font", v)}
            options={fontOptions}
          />
        </Field>
        <Field label="Font Size">
          <Select
            value={String(params.fontSize)}
            onChange={(v) => update("fontSize", Number(v))}
            options={fontSizeOptions}
          />
        </Field>
        <Field label="Font Weight">
          <Select
            value={params.weight}
            onChange={(v) => update("weight", v as Weight)}
            options={WEIGHT_OPTIONS.map((w) => ({ value: w, label: w }))}
          />
        </Field>
        <Field label="Font Color">
          <input
            type="color"
            value={rgbStringToHex(params.color)}
            onChange={(e) => update("color", hexToRgbString(e.target.value))}
            className="w-full h-6 px-1 border border-slate-200 rounded-md cursor-pointer"
          />
        </Field>
      </Section>

      {/* ── Text Alignment ── */}
      <Section label="Text Alignment">
        <Field label="Horizontal Alignment">
          <SegmentedControl
            value={params.align}
            onChange={(v) => update("align", v as Align)}
            options={horizontals}
          />
        </Field>
        <Field label="Vertical Alignment">
          <SegmentedControl
            value={params.verticalAlignment}
            onChange={(v) => update("verticalAlignment", v as VerticalAlignment)}
            options={verticals}
          />
        </Field>
      </Section>

      {/* ── Input ── */}
      <Section label="Input">
        <Field label="Max Length">
          <NumberInput
            value={params.maxLength}
            min={0}
            placeholder="0 = unlimited"
            onChange={(v) => update("maxLength", v)}
          />
        </Field>
        <Field label="Placeholder">
          <TextInput
            value={params.placeholder}
            placeholder="Placeholder text"
            onChange={(v) => update("placeholder", v)}
          />
        </Field>
        <Field label="Default Value">
          <TextInput
            value={params.defaultValue}
            placeholder="Default text"
            onChange={(v) => update("defaultValue", v)}
          />
        </Field>
        <Field label="Input Restriction">
          <Select
            value={params.inputRestriction}
            onChange={(v) => update("inputRestriction", v as InputRestriction)}
            options={restrictionOptions}
          />
        </Field>
      </Section>

      {/* ── Behavior ── */}
      <Section label="Behavior">
        <Toggle
          label="Required"
          checked={params.required}
          onChange={(v) => update("required", v)}
        />
        <Toggle
          label="Read Only"
          checked={params.readOnly}
          onChange={(v) => update("readOnly", v)}
        />
        <Toggle
          label="Hidden"
          checked={params.hidden}
          onChange={(v) => update("hidden", v)}
        />
      </Section>
    </div>
  );
}
