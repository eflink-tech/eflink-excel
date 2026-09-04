export type Rgba = { r: number; g: number; b: number; a: number };
export type Hsv = { h: number; s: number; v: number };

const HEX3_RE = /^#([0-9a-fA-F]{3})$/;
const HEX6_RE = /^#([0-9a-fA-F]{6})$/;
const RGBA_RE =
  /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*([\d.]+))?\s*\)$/;

function clampByte(value: number): number {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function clampAlpha(value: number): number {
  return Math.max(0, Math.min(1, value));
}

function expandHex3(hex: string): string {
  return hex
    .split('')
    .map((ch) => ch + ch)
    .join('');
}

function parseHexChannels(hex: string): { r: number; g: number; b: number } | null {
  const match3 = HEX3_RE.exec(hex);
  if (match3) {
    const expanded = expandHex3(match3[1]!);
    return {
      r: Number.parseInt(expanded.slice(0, 2), 16),
      g: Number.parseInt(expanded.slice(2, 4), 16),
      b: Number.parseInt(expanded.slice(4, 6), 16),
    };
  }

  const match6 = HEX6_RE.exec(hex);
  if (match6) {
    return {
      r: Number.parseInt(match6[1]!.slice(0, 2), 16),
      g: Number.parseInt(match6[1]!.slice(2, 4), 16),
      b: Number.parseInt(match6[1]!.slice(4, 6), 16),
    };
  }

  return null;
}

export function normalizeColor(input: string): string | null {
  const trimmed = input.trim();
  if (!trimmed) return null;

  const hexMatch = trimmed.match(/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/);
  if (hexMatch) {
    const channels = parseHexChannels(trimmed);
    if (!channels) return null;
    const { r, g, b } = channels;
    return formatCssColor(r, g, b, 1);
  }

  const rgbaMatch = RGBA_RE.exec(trimmed);
  if (rgbaMatch) {
    const r = clampByte(Number(rgbaMatch[1]));
    const g = clampByte(Number(rgbaMatch[2]));
    const b = clampByte(Number(rgbaMatch[3]));
    const a = rgbaMatch[4] !== undefined ? clampAlpha(Number(rgbaMatch[4])) : 1;
    return formatCssColor(r, g, b, a);
  }

  return null;
}

export function toRgba(color: string): Rgba | null {
  const normalized = normalizeColor(color);
  if (!normalized) return null;

  if (normalized.startsWith('#')) {
    const channels = parseHexChannels(normalized);
    if (!channels) return null;
    return { ...channels, a: 1 };
  }

  const rgbaMatch = RGBA_RE.exec(normalized);
  if (!rgbaMatch) return null;

  return {
    r: clampByte(Number(rgbaMatch[1])),
    g: clampByte(Number(rgbaMatch[2])),
    b: clampByte(Number(rgbaMatch[3])),
    a: rgbaMatch[4] !== undefined ? clampAlpha(Number(rgbaMatch[4])) : 1,
  };
}

export function formatCssColor(r: number, g: number, b: number, a: number): string {
  const rr = clampByte(r).toString(16).padStart(2, '0');
  const gg = clampByte(g).toString(16).padStart(2, '0');
  const bb = clampByte(b).toString(16).padStart(2, '0');

  if (a === 1) {
    return `#${rr}${gg}${bb}`;
  }

  return `rgba(${clampByte(r)}, ${clampByte(g)}, ${clampByte(b)}, ${clampAlpha(a)})`;
}

export function rgbToHsv(r: number, g: number, b: number): Hsv {
  const rn = clampByte(r) / 255;
  const gn = clampByte(g) / 255;
  const bn = clampByte(b) / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const delta = max - min;

  let h = 0;
  const v = max * 100;
  const s = max === 0 ? 0 : (delta / max) * 100;

  if (delta !== 0) {
    if (max === rn) {
      h = ((gn - bn) / delta + (gn < bn ? 6 : 0)) * 60;
    } else if (max === gn) {
      h = ((bn - rn) / delta + 2) * 60;
    } else {
      h = ((rn - gn) / delta + 4) * 60;
    }
  }

  return { h, s, v };
}

export function hsvToRgb(h: number, s: number, v: number): { r: number; g: number; b: number } {
  const sn = Math.max(0, Math.min(100, s)) / 100;
  const vn = Math.max(0, Math.min(100, v)) / 100;
  const c = vn * sn;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = vn - c;

  let rp = 0;
  let gp = 0;
  let bp = 0;
  const hn = ((h % 360) + 360) % 360;

  if (hn < 60) {
    rp = c;
    gp = x;
  } else if (hn < 120) {
    rp = x;
    gp = c;
  } else if (hn < 180) {
    gp = c;
    bp = x;
  } else if (hn < 240) {
    gp = x;
    bp = c;
  } else if (hn < 300) {
    rp = x;
    bp = c;
  } else {
    rp = c;
    bp = x;
  }

  return {
    r: clampByte((rp + m) * 255),
    g: clampByte((gp + m) * 255),
    b: clampByte((bp + m) * 255),
  };
}
