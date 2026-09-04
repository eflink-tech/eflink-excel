import { useCallback, useEffect, useRef, useState } from 'react';
import {
  formatCssColor,
  hsvToRgb,
  rgbToHsv,
  toRgba,
  type Hsv,
  type Rgba,
} from './colorFormat';

export type ColorAdvancedPanelProps = {
  value?: string | null;
  onChange: (color: string) => void;
};

const DEFAULT_RGBA: Rgba = { r: 51, g: 51, b: 51, a: 1 };

function parseValue(value?: string | null): Rgba {
  if (!value) return DEFAULT_RGBA;
  return toRgba(value) ?? DEFAULT_RGBA;
}

function rgbaToState(rgba: Rgba): { hsv: Hsv; aPercent: number } {
  const hsv = rgbToHsv(rgba.r, rgba.g, rgba.b);
  return { hsv, aPercent: Math.round(rgba.a * 100) };
}

function stateToCss(hsv: Hsv, aPercent: number): string {
  const { r, g, b } = hsvToRgb(hsv.h, hsv.s, hsv.v);
  return formatCssColor(r, g, b, aPercent / 100);
}

function clampPercent(value: number): number {
  return Math.max(0, Math.min(100, Math.round(value)));
}

function parseHexInput(raw: string): { r: number; g: number; b: number } | null {
  const hex = raw.replace(/^#/, '').trim();
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return null;
  return {
    r: Number.parseInt(hex.slice(0, 2), 16),
    g: Number.parseInt(hex.slice(2, 4), 16),
    b: Number.parseInt(hex.slice(4, 6), 16),
  };
}

const CHECKERBOARD =
  'linear-gradient(45deg, #ccc 25%, transparent 25%), linear-gradient(-45deg, #ccc 25%, transparent 25%), linear-gradient(45deg, transparent 75%, #ccc 75%), linear-gradient(-45deg, transparent 75%, #ccc 75%)';

export function ColorAdvancedPanel({ value, onChange }: ColorAdvancedPanelProps) {
  const initial = rgbaToState(parseValue(value));
  const [hsv, setHsv] = useState<Hsv>(initial.hsv);
  const [aPercent, setAPercent] = useState(initial.aPercent);
  const svRef = useRef<HTMLDivElement>(null);
  const hueRef = useRef<HTMLDivElement>(null);
  const alphaRef = useRef<HTMLDivElement>(null);
  const draggingSv = useRef(false);
  const draggingHue = useRef(false);
  const draggingAlpha = useRef(false);
  const hsvRef = useRef(hsv);
  const aPercentRef = useRef(aPercent);
  const onChangeRef = useRef(onChange);

  hsvRef.current = hsv;
  aPercentRef.current = aPercent;
  onChangeRef.current = onChange;

  useEffect(() => {
    const next = rgbaToState(parseValue(value));
    setHsv(next.hsv);
    setAPercent(next.aPercent);
  }, [value]);

  const emitChange = useCallback((nextHsv: Hsv, nextAPercent: number) => {
    onChangeRef.current(stateToCss(nextHsv, nextAPercent));
  }, []);

  const updateFromSv = useCallback((clientX: number, clientY: number) => {
    const el = svRef.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const x = Math.max(0, Math.min(rect.width, clientX - rect.left));
    const y = Math.max(0, Math.min(rect.height, clientY - rect.top));
    const s = (x / rect.width) * 100;
    const v = 100 - (y / rect.height) * 100;
    const next = { ...hsvRef.current, s, v };
    setHsv(next);
    emitChange(next, aPercentRef.current);
  }, [emitChange]);

  const updateFromHue = useCallback(
    (clientX: number) => {
      const el = hueRef.current;
      if (!el) return;
      const rect = el.getBoundingClientRect();
      const x = Math.max(0, Math.min(rect.width, clientX - rect.left));
      const h = (x / rect.width) * 360;
      const next = { ...hsvRef.current, h };
      setHsv(next);
      emitChange(next, aPercentRef.current);
    },
    [emitChange],
  );

  const updateFromAlpha = useCallback(
    (clientX: number) => {
      const el = alphaRef.current;
      if (!el) return;
      const rect = el.getBoundingClientRect();
      const x = Math.max(0, Math.min(rect.width, clientX - rect.left));
      const next = clampPercent((x / rect.width) * 100);
      setAPercent(next);
      emitChange(hsvRef.current, next);
    },
    [emitChange],
  );

  const { r, g, b } = hsvToRgb(hsv.h, hsv.s, hsv.v);
  const opaqueColor = formatCssColor(r, g, b, 1);
  const previewColor = formatCssColor(r, g, b, aPercent / 100);
  const pureHue = hsvToRgb(hsv.h, 100, 100);
  const hueColor = formatCssColor(pureHue.r, pureHue.g, pureHue.b, 1);

  const handleHexChange = (raw: string) => {
    const parsed = parseHexInput(raw);
    if (!parsed) return;
    const nextHsv = rgbToHsv(parsed.r, parsed.g, parsed.b);
    setHsv(nextHsv);
    emitChange(nextHsv, aPercent);
  };

  const handleChannelChange = (channel: 'r' | 'g' | 'b', raw: string) => {
    const num = Number.parseInt(raw, 10);
    if (Number.isNaN(num)) return;
    const clamped = Math.max(0, Math.min(255, num));
    const nextRgb = { r, g, b, [channel]: clamped };
    const nextHsv = rgbToHsv(nextRgb.r, nextRgb.g, nextRgb.b);
    setHsv(nextHsv);
    emitChange(nextHsv, aPercent);
  };

  const handleAlphaChange = (raw: string) => {
    const num = Number.parseInt(raw, 10);
    if (Number.isNaN(num)) return;
    const next = clampPercent(num);
    setAPercent(next);
    emitChange(hsv, next);
  };

  useEffect(() => {
    const onMove = (event: PointerEvent) => {
      if (draggingSv.current) updateFromSv(event.clientX, event.clientY);
      if (draggingHue.current) updateFromHue(event.clientX);
      if (draggingAlpha.current) updateFromAlpha(event.clientX);
    };
    const onUp = () => {
      draggingSv.current = false;
      draggingHue.current = false;
      draggingAlpha.current = false;
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
    return () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
    };
  }, [updateFromAlpha, updateFromHue, updateFromSv]);

  return (
    <div className="w-[280px] rounded-lg border border-[#e7e9eb] bg-white p-3 shadow-[0_4px_16px_rgba(0,0,0,0.12)]">
      <div className="flex gap-3">
        <div className="flex min-w-0 flex-1 flex-col gap-2">
          <div
            ref={svRef}
            data-testid="sv-area"
            className="relative h-[140px] cursor-crosshair rounded border border-[#e7e9eb]"
            style={{
              backgroundColor: hueColor,
              backgroundImage:
                'linear-gradient(to right, #fff, transparent), linear-gradient(to top, #000, transparent)',
            }}
            onPointerDown={(event) => {
              draggingSv.current = true;
              updateFromSv(event.clientX, event.clientY);
            }}
          >
            <span
              className="pointer-events-none absolute h-3 w-3 -translate-x-1/2 -translate-y-1/2 rounded-full border-2 border-white shadow-[0_0_0_1px_rgba(0,0,0,0.3)]"
              style={{
                left: `${hsv.s}%`,
                top: `${100 - hsv.v}%`,
              }}
            />
          </div>

          <div
            ref={hueRef}
            data-testid="hue-slider"
            className="relative h-3 cursor-pointer rounded border border-[#e7e9eb]"
            style={{
              background:
                'linear-gradient(to right, #ff0000, #ffff00, #00ff00, #00ffff, #0000ff, #ff00ff, #ff0000)',
            }}
            onPointerDown={(event) => {
              draggingHue.current = true;
              updateFromHue(event.clientX);
            }}
          >
            <span
              className="pointer-events-none absolute top-1/2 h-3 w-1 -translate-x-1/2 -translate-y-1/2 rounded-sm border border-[#1f2329] bg-white"
              style={{ left: `${(hsv.h / 360) * 100}%` }}
            />
          </div>

          <div
            ref={alphaRef}
            data-testid="alpha-slider"
            className="relative h-3 cursor-pointer overflow-hidden rounded border border-[#e7e9eb]"
            style={{
              backgroundColor: '#fff',
              backgroundImage: `${CHECKERBOARD}`,
              backgroundSize: '8px 8px',
              backgroundPosition: '0 0, 0 4px, 4px -4px, -4px 0',
            }}
            onPointerDown={(event) => {
              draggingAlpha.current = true;
              updateFromAlpha(event.clientX);
            }}
          >
            <div
              className="absolute inset-0"
              style={{
                background: `linear-gradient(to right, transparent, ${opaqueColor})`,
              }}
            />
            <span
              className="pointer-events-none absolute top-1/2 h-3 w-1 -translate-x-1/2 -translate-y-1/2 rounded-sm border border-[#1f2329] bg-white"
              style={{ left: `${aPercent}%` }}
            />
          </div>
        </div>

        <div
          data-testid="color-preview"
          className="h-10 w-10 shrink-0 rounded border border-[#e7e9eb]"
          style={{
            backgroundColor: previewColor,
            backgroundImage: aPercent < 100 ? CHECKERBOARD : undefined,
            backgroundSize: aPercent < 100 ? '8px 8px' : undefined,
          }}
        />
      </div>

      <div className="mt-3 grid grid-cols-5 gap-2">
        <label className="flex flex-col gap-1 text-center">
          <input
            aria-label="Hex"
            value={formatCssColor(r, g, b, 1).slice(1).toUpperCase()}
            onChange={(event) => handleHexChange(event.target.value)}
            className="w-full rounded border border-[#e7e9eb] px-1 py-1 text-center text-[12px] text-[#1f2329] outline-none focus:border-[#3370ff]"
          />
          <span className="text-[11px] text-[#8f959e]">Hex</span>
        </label>
        <label className="flex flex-col gap-1 text-center">
          <input
            aria-label="R"
            type="number"
            min={0}
            max={255}
            value={r}
            onChange={(event) => handleChannelChange('r', event.target.value)}
            className="w-full rounded border border-[#e7e9eb] px-1 py-1 text-center text-[12px] text-[#1f2329] outline-none focus:border-[#3370ff]"
          />
          <span className="text-[11px] text-[#8f959e]">R</span>
        </label>
        <label className="flex flex-col gap-1 text-center">
          <input
            aria-label="G"
            type="number"
            min={0}
            max={255}
            value={g}
            onChange={(event) => handleChannelChange('g', event.target.value)}
            className="w-full rounded border border-[#e7e9eb] px-1 py-1 text-center text-[12px] text-[#1f2329] outline-none focus:border-[#3370ff]"
          />
          <span className="text-[11px] text-[#8f959e]">G</span>
        </label>
        <label className="flex flex-col gap-1 text-center">
          <input
            aria-label="B"
            type="number"
            min={0}
            max={255}
            value={b}
            onChange={(event) => handleChannelChange('b', event.target.value)}
            className="w-full rounded border border-[#e7e9eb] px-1 py-1 text-center text-[12px] text-[#1f2329] outline-none focus:border-[#3370ff]"
          />
          <span className="text-[11px] text-[#8f959e]">B</span>
        </label>
        <label className="flex flex-col gap-1 text-center">
          <input
            aria-label="A"
            type="number"
            min={0}
            max={100}
            value={aPercent}
            onChange={(event) => handleAlphaChange(event.target.value)}
            className="w-full rounded border border-[#e7e9eb] px-1 py-1 text-center text-[12px] text-[#1f2329] outline-none focus:border-[#3370ff]"
          />
          <span className="text-[11px] text-[#8f959e]">A</span>
        </label>
      </div>
    </div>
  );
}
