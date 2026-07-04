import fs from "node:fs";
import path from "node:path";
import { DefaultResourceLoader } from "@earendil-works/pi-coding-agent";

type DefaultResourceLoaderInit = ConstructorParameters<typeof DefaultResourceLoader>[0];

export const EMBEDDED_PI_RESOURCE_LOADER_DISCOVERY_OPTIONS = {
  noExtensions: true,
  noSkills: true,
  noPromptTemplates: true,
  noThemes: true,
  noContextFiles: true,
} satisfies Partial<DefaultResourceLoaderInit>;

export function createEmbeddedPiResourceLoader(
  options: Pick<
    DefaultResourceLoaderInit,
    "cwd" | "agentDir" | "settingsManager" | "extensionFactories"
  >,
): DefaultResourceLoader {
  return new DefaultResourceLoader({
    ...options,
    ...EMBEDDED_PI_RESOURCE_LOADER_DISCOVERY_OPTIONS,
  });
}

/**
 * Mtime-tracked cache for resource loader reloads.
 * Avoids re-reading settings files from disk on every embedded run
 * when nothing has changed.
 */
type ResourceLoaderCacheEntry = {
  lastMtimeMs: number;
  lastCheckMs: number;
};

const resourceLoaderCache = new Map<string, ResourceLoaderCacheEntry>();
const RESOURCE_LOADER_CACHE_MAX_SIZE = 64; // prevent unbounded growth across workspaces
const RESOURCE_LOADER_CACHE_TTL_MS = 30_000; // revalidate at most every 30s
const RESOURCE_LOADER_STALE_MS = 2_000; // skip reload when mtime changed within 2s

function getSettingsMtime(cwd: string, agentDir?: string): number {
  let maxMtime = 0;
  const dirs = [cwd];
  if (agentDir) {
    dirs.push(agentDir);
  }
  for (const dir of dirs) {
    try {
      const settingsPath = path.join(dir, ".pi", "settings.json");
      const stat = fs.statSync(settingsPath, { throwIfNoEntry: false });
      if (stat) {
        maxMtime = Math.max(maxMtime, stat.mtimeMs);
      }
    } catch {
      // settings file may not exist
    }
  }
  return maxMtime;
}

/**
 * Returns true if the resource loader should skip reload() because
 * settings files haven't changed since the last cache check.
 */
export function shouldSkipResourceLoaderReload(cwd: string, agentDir?: string): boolean {
  const cacheKey = `${cwd}\x00${agentDir ?? ""}`;
  const now = Date.now();
  const entry = resourceLoaderCache.get(cacheKey);
  const currentMtime = getSettingsMtime(cwd, agentDir);

  if (entry) {
    // If mtime hasn't changed and cache is fresh enough, skip
    if (
      entry.lastMtimeMs === currentMtime &&
      now - entry.lastCheckMs < RESOURCE_LOADER_CACHE_TTL_MS
    ) {
      return true;
    }
    // If mtime changed recently, also skip — prevents reload storms when
    // settings files are being actively written (e.g. by a simultaneous run).
    if (entry.lastMtimeMs !== currentMtime && now - entry.lastCheckMs < RESOURCE_LOADER_STALE_MS) {
      return true;
    }
  }

  resourceLoaderCache.set(cacheKey, {
    lastMtimeMs: currentMtime,
    lastCheckMs: now,
  });
  // Evict oldest entry when cache exceeds max size.
  if (resourceLoaderCache.size > RESOURCE_LOADER_CACHE_MAX_SIZE) {
    const oldest = resourceLoaderCache.keys().next().value;
    if (oldest) {
      resourceLoaderCache.delete(oldest);
    }
  }
  return false;
}

export function clearResourceLoaderCacheForTest(): void {
  resourceLoaderCache.clear();
}
