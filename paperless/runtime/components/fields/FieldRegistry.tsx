"use client";

import type { ComponentType } from "react";
import type { OverlayModel, OverlayConfig } from "@/types/overlay";

/**
 * Standard props passed to every registered field component.
 * Field components are presentational: they receive value + onChange.
 * All state management is handled by RuntimeField via the Zustand store.
 */
export interface FieldComponentProps {
  overlay: OverlayModel;
  value: string | boolean | null;
  onChange: (value: string | boolean | null) => void;
  onBlur?: () => void;
  production?: boolean;
  config?: OverlayConfig;
  disabled?: boolean;
  readOnly?: boolean;
  required?: boolean;
  placeholder?: string;
}

export interface FieldRegistration {
  type: string;
  component: ComponentType<FieldComponentProps>;
}

/**
 * FieldRegistry — the single source of truth for mapping overlay/types
 * to React field components.
 *
 * Instead of a switch-case in RuntimeField, components are registered
 * here at module load time. Registering a new field type requires:
 *   1. Create the component
 *   2. Import and register it here
 *   3. No changes to RuntimeField, RuntimeCanvas, or any dispatcher
 */
const registry = new Map<string, ComponentType<FieldComponentProps>>();

export function registerField(type: string, component: ComponentType<FieldComponentProps>) {
  registry.set(type, component);
}

export function getFieldComponent(type: string): ComponentType<FieldComponentProps> | null {
  return registry.get(type) ?? null;
}

export function getRegisteredTypes(): string[] {
  return Array.from(registry.keys());
}

export { registry as _registry };
