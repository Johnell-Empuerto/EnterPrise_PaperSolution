"use client";

import { useState, useEffect, useCallback } from "react";
import type { RuntimeForm } from "@/types/runtime";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:5090";

interface UseRuntimeResult {
  /** The loaded runtime form, or null if not yet loaded. */
  runtimeForm: RuntimeForm | null;
  /** Whether the request is in flight. */
  loading: boolean;
  /** Error message if the request failed. */
  error: string | null;
  /** Manually trigger a reload (re-fetches current templateId). */
  reload: () => void;
  /**
   * Programmatically load a specific template ID.
   * This bypasses the reactive templateId prop for the upload flow.
   */
  loadByTemplateId: (templateId: string) => Promise<void>;
}

/**
 * Fetches the RuntimeForm from GET /api/runtime/{templateId}.
 * Supports both reactive (templateId prop) and programmatic (loadByTemplateId) modes.
 */
export function useRuntime(templateId: string | null): UseRuntimeResult {
  const [runtimeForm, setRuntimeForm] = useState<RuntimeForm | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [reloadCounter, setReloadCounter] = useState(0);

  // Track the most recently loaded templateId so loadByTemplateId can
  // persist across re-renders while the parent also changes the prop.
  const [loadedTemplateId, setLoadedTemplateId] = useState<string | null>(null);

  const doFetch = useCallback(async (id: string) => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`${API_BASE_URL}/api/form/runtime/${encodeURIComponent(id)}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}: ${res.statusText}`);
      const json = await res.json();
      if (json.success && json.data) {
        setRuntimeForm(json.data as RuntimeForm);
      } else {
        throw new Error(json.message ?? "Unknown error");
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load runtime");
    } finally {
      setLoading(false);
    }
  }, []);

  const reload = useCallback(() => {
    setReloadCounter((c) => c + 1);
  }, []);

  const loadByTemplateId = useCallback(
    async (id: string) => {
      setLoadedTemplateId(id);
      await doFetch(id);
    },
    [doFetch]
  );

  // Reactive fetch: when templateId prop changes, fetch automatically.
  useEffect(() => {
    if (!templateId) {
      // Only clear if no programmatic load is active
      if (!loadedTemplateId) {
        setRuntimeForm(null);
        setLoading(false);
        setError(null);
      }
      return;
    }

    let cancelled = false;

    // If the reactive templateId differs from the programmatic one,
    // use the reactive one (parent change takes priority).
    setLoadedTemplateId(templateId);
    setLoading(true);
    setError(null);

    fetch(`${API_BASE_URL}/api/form/runtime/${encodeURIComponent(templateId)}`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP ${res.status}: ${res.statusText}`);
        return res.json();
      })
      .then((json) => {
        if (cancelled) return;
        if (json.success && json.data) {
          setRuntimeForm(json.data as RuntimeForm);
        } else {
          throw new Error(json.message ?? "Unknown error");
        }
      })
      .catch((err) => {
        if (cancelled) return;
        setError(err instanceof Error ? err.message : "Failed to load runtime");
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });

    return () => {
      cancelled = true;
    };
  }, [templateId, reloadCounter, doFetch, loadedTemplateId]);

  return { runtimeForm, loading, error, reload, loadByTemplateId };
}
