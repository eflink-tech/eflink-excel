import { ChevronRight } from 'lucide-react';
import { STANDARD_COLORS, THEME_COLORS } from './colorPalette';
import { RECENT_MAX } from './recentColors';

export type ColorPalettePanelProps = {
  recent: string[];
  onSelect: (color: string) => void;
  onDefault: () => void;
  onMore: () => void;
};

function themeSwatches(): string[] {
  const swatches: string[] = [];
  for (let row = 0; row < 6; row += 1) {
    for (let col = 0; col < THEME_COLORS.length; col += 1) {
      swatches.push(THEME_COLORS[col]![row]!);
    }
  }
  return swatches;
}

const THEME_SWATCHES = themeSwatches();

function ColorSwatch({
  color,
  testId,
  title,
  filled,
  onClick,
}: {
  color: string;
  testId: string;
  title?: string;
  filled?: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      data-testid={testId}
      data-filled={filled === undefined ? undefined : filled ? 'true' : 'false'}
      title={title ?? color}
      aria-label={title ?? color}
      onClick={onClick}
      className="h-[18px] w-[18px] rounded-[2px] border border-[#e7e9eb] transition-transform hover:scale-110 hover:border-[#8f959e]"
      style={{ backgroundColor: color }}
    />
  );
}

export function ColorPalettePanel({ recent, onSelect, onDefault, onMore }: ColorPalettePanelProps) {
  const recentSlots = Array.from({ length: RECENT_MAX }, (_, index) => recent[index] ?? null);

  return (
    <div className="w-[240px] rounded-lg border border-[#e7e9eb] bg-white p-3 shadow-[0_4px_16px_rgba(0,0,0,0.12)]">
      <button
        type="button"
        onClick={onDefault}
        className="mb-3 w-full rounded border border-[#e7e9eb] bg-[#f2f3f4] px-3 py-2 text-[13px] text-[#1f2329] hover:bg-[#e7e9eb]"
      >
        默认
      </button>

      <div className="grid grid-cols-10 gap-[3px]">
        {THEME_SWATCHES.map((color, index) => (
          <ColorSwatch
            key={`theme-${index}`}
            color={color}
            testId="theme-color"
            onClick={() => onSelect(color)}
          />
        ))}
      </div>

      <div className="mt-3 text-[11px] text-[#8f959e]">标准色</div>
      <div className="mt-1 grid grid-cols-10 gap-[3px]">
        {STANDARD_COLORS.map((color) => (
          <ColorSwatch
            key={`standard-${color}`}
            color={color}
            testId="standard-color"
            onClick={() => onSelect(color)}
          />
        ))}
      </div>

      <div className="mt-3 text-[11px] text-[#8f959e]">最近使用</div>
      <div className="mt-1 grid grid-cols-10 gap-[3px]">
        {recentSlots.map((color, index) =>
          color ? (
            <ColorSwatch
              key={`recent-${color}-${index}`}
              color={color}
              testId="recent-slot"
              filled
              onClick={() => onSelect(color)}
            />
          ) : (
            <button
              key={`recent-empty-${index}`}
              type="button"
              data-testid="recent-slot"
              data-filled="false"
              aria-label="空最近色槽"
              disabled
              className="h-[18px] w-[18px] rounded-[2px] border border-[#e7e9eb] bg-white"
            />
          ),
        )}
      </div>

      <div className="my-3 border-t border-[#e7e9eb]" />

      <button
        type="button"
        onClick={onMore}
        className="flex min-h-[36px] w-full items-center gap-2 rounded px-3 py-2 text-left text-[13px] text-[#1f2329] hover:bg-[#f2f3f4]"
      >
        <span
          aria-hidden
          className="h-4 w-4 shrink-0 rounded-full border border-[#e7e9eb]"
          style={{
            background:
              'conic-gradient(red, yellow, lime, aqua, blue, magenta, red)',
          }}
        />
        <span className="flex-1">更多颜色</span>
        <ChevronRight size={14} className="text-[#8f959e]" />
      </button>
    </div>
  );
}
