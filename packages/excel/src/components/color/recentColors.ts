import { normalizeColor } from './colorFormat';

export const RECENT_MAX = 10;

function readRaw(storageKey: string): unknown {
  try {
    const raw = localStorage.getItem(storageKey);
    if (!raw) return [];
    return JSON.parse(raw);
  } catch {
    return [];
  }
}

function writeRaw(storageKey: string, colors: string[]): void {
  localStorage.setItem(storageKey, JSON.stringify(colors));
}

export function loadRecent(storageKey: string): string[] {
  const raw = readRaw(storageKey);
  if (!Array.isArray(raw)) return [];

  const seen = new Set<string>();
  const result: string[] = [];

  for (const item of raw) {
    if (typeof item !== 'string') continue;
    const normalized = normalizeColor(item);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    result.push(normalized);
    if (result.length >= RECENT_MAX) break;
  }

  return result;
}

export function pushRecent(storageKey: string, color: string): string[] {
  const normalized = normalizeColor(color);
  if (!normalized) {
    return loadRecent(storageKey);
  }

  const current = loadRecent(storageKey);
  const next = [normalized, ...current.filter((c) => c !== normalized)].slice(0, RECENT_MAX);
  writeRaw(storageKey, next);
  return next;
}
