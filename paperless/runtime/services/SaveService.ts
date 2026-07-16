import type { RuntimeStoreApi } from "../store";

export interface SaveQueueEntry {
  fieldId: string;
  value: string | boolean | null;
  timestamp: number;
}

export interface SaveResult {
  success: boolean;
  error?: string;
}

export interface SaveServiceOptions {
  debounceMs: number;
  onSave: (values: Record<string, string | boolean | null>) => Promise<SaveResult>;
  maxRetries: number;
}

const DEFAULT_OPTIONS: SaveServiceOptions = {
  debounceMs: 2000,
  onSave: async () => ({ success: true }),
  maxRetries: 3,
};

/**
 * Manages auto-save with debounce, batching, and retry logic.
 * Accumulates changed values and flushes them after a debounce period.
 */
export class SaveService {
  private options: SaveServiceOptions;
  private queue: Map<string, SaveQueueEntry> = new Map();
  private timer: ReturnType<typeof setTimeout> | null = null;
  private retryCount = 0;
  private isSaving = false;

  constructor(options?: Partial<SaveServiceOptions>) {
    this.options = { ...DEFAULT_OPTIONS, ...options };
  }

  enqueue(fieldId: string, value: string | boolean | null) {
    this.queue.set(fieldId, {
      fieldId,
      value,
      timestamp: Date.now(),
    });
    this.scheduleFlush();
  }

  enqueueMultiple(entries: Array<{ fieldId: string; value: string | boolean | null }>) {
    for (const entry of entries) {
      this.queue.set(entry.fieldId, {
        ...entry,
        timestamp: Date.now(),
      });
    }
    this.scheduleFlush();
  }

  async flush(): Promise<SaveResult> {
    if (this.queue.size === 0) return { success: true };
    if (this.isSaving) return { success: false, error: "Already saving" };

    this.cancelTimer();
    this.isSaving = true;

    const batch = new Map(this.queue);
    this.queue.clear();

    try {
      const values: Record<string, string | boolean | null> = {};
      for (const [, entry] of batch) {
        values[entry.fieldId] = entry.value;
      }

      const result = await this.options.onSave(values);

      if (result.success) {
        this.retryCount = 0;
        return result;
      }

      if (this.retryCount < this.options.maxRetries) {
        this.retryCount++;
        for (const [fieldId, entry] of batch) {
          this.queue.set(fieldId, entry);
        }
        this.scheduleFlush();
      }

      return result;
    } finally {
      this.isSaving = false;
    }
  }

  get pendingCount(): number {
    return this.queue.size;
  }

  get isBusy(): boolean {
    return this.isSaving;
  }

  destroy() {
    this.cancelTimer();
    this.queue.clear();
  }

  private scheduleFlush() {
    this.cancelTimer();
    this.timer = setTimeout(() => {
      this.flush();
    }, this.options.debounceMs);
  }

  private cancelTimer() {
    if (this.timer !== null) {
      clearTimeout(this.timer);
      this.timer = null;
    }
  }
}
