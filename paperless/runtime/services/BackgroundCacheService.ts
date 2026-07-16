export interface CachedImage {
  url: string;
  blob: Blob;
  objectUrl: string;
  loadedAt: number;
  accessCount: number;
}

export interface BackgroundCacheOptions {
  maxEntries: number;
}

const DEFAULT_OPTIONS: BackgroundCacheOptions = {
  maxEntries: 10,
};

/**
 * LRU cache for background images.
 * Prevents memory exhaustion by limiting the number of cached images.
 * Automatically revokes object URLs when evicting entries.
 */
export class BackgroundCacheService {
  private cache: Map<string, CachedImage> = new Map();
  private options: BackgroundCacheOptions;

  constructor(options?: Partial<BackgroundCacheOptions>) {
    this.options = { ...DEFAULT_OPTIONS, ...options };
  }

  async getOrLoad(url: string): Promise<string> {
    const existing = this.cache.get(url);
    if (existing) {
      existing.accessCount++;
      existing.loadedAt = Date.now();
      return existing.objectUrl;
    }

    const blob = await this.fetchImage(url);
    const objectUrl = URL.createObjectURL(blob);

    this.add(url, blob, objectUrl);
    return objectUrl;
  }

  has(url: string): boolean {
    return this.cache.has(url);
  }

  get(url: string): string | null {
    const entry = this.cache.get(url);
    if (!entry) return null;
    entry.accessCount++;
    return entry.objectUrl;
  }

  invalidate(url: string) {
    const entry = this.cache.get(url);
    if (entry) {
      URL.revokeObjectURL(entry.objectUrl);
      this.cache.delete(url);
    }
  }

  clear() {
    for (const [, entry] of this.cache) {
      URL.revokeObjectURL(entry.objectUrl);
    }
    this.cache.clear();
  }

  get size(): number {
    return this.cache.size;
  }

  destroy() {
    this.clear();
  }

  private async fetchImage(url: string): Promise<Blob> {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`Failed to load background image: ${response.status}`);
    }
    return response.blob();
  }

  private add(url: string, blob: Blob, objectUrl: string) {
    if (this.cache.size >= this.options.maxEntries) {
      this.evictLeastRecentlyUsed();
    }

    this.cache.set(url, {
      url,
      blob,
      objectUrl,
      loadedAt: Date.now(),
      accessCount: 1,
    });
  }

  private evictLeastRecentlyUsed() {
    let oldest: string | null = null;
    let oldestTime = Infinity;

    for (const [url, entry] of this.cache) {
      if (entry.loadedAt < oldestTime) {
        oldestTime = entry.loadedAt;
        oldest = url;
      }
    }

    if (oldest) {
      this.invalidate(oldest);
    }
  }
}
