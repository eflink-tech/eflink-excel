import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { ColorAdvancedPanel } from './ColorAdvancedPanel';
import { ColorPalettePanel } from './ColorPalettePanel';
import { loadRecent, pushRecent } from './recentColors';
import { pickAdvancedSide, type PanelSide } from './pickAdvancedSide';

export type ColorSelectOptions = {
  /** When false, apply color without closing the popover (advanced panel preview). Default true. */
  close?: boolean;
};

export type StandaloneColorPickerProps = {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  value?: string | null;
  recentStorageKey: string;
  onSelect: (color: string, options?: ColorSelectOptions) => void;
  onDefault: () => void;
  className?: string;
};

export function ColorPicker({
  open,
  onOpenChange: _onOpenChange,
  value,
  recentStorageKey,
  onSelect,
  onDefault,
  className,
}: StandaloneColorPickerProps) {
  const [recent, setRecent] = useState<string[]>([]);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [advancedSide, setAdvancedSide] = useState<PanelSide>('right');
  const paletteAnchorRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (open) {
      setRecent(loadRecent(recentStorageKey));
    } else {
      setShowAdvanced(false);
      setAdvancedSide('right');
    }
  }, [open, recentStorageKey]);

  useLayoutEffect(() => {
    if (!showAdvanced) return;

    const updateSide = () => {
      const el = paletteAnchorRef.current;
      if (!el) return;
      const rect = el.getBoundingClientRect();
      setAdvancedSide(
        pickAdvancedSide({
          paletteLeft: rect.left,
          paletteRight: rect.right,
          viewportWidth: window.innerWidth,
        }),
      );
    };

    updateSide();
    window.addEventListener('resize', updateSide);
    return () => window.removeEventListener('resize', updateSide);
  }, [showAdvanced]);

  const handlePaletteSelect = (color: string) => {
    setRecent(pushRecent(recentStorageKey, color));
    onSelect(color);
  };

  const handleAdvancedChange = (color: string) => {
    setRecent(pushRecent(recentStorageKey, color));
    onSelect(color, { close: false });
  };

  const handleDefault = () => {
    onDefault();
  };

  if (!open) return null;

  return (
    <div
      data-testid="color-picker"
      data-advanced-side={showAdvanced ? advancedSide : undefined}
      className={['relative', className].filter(Boolean).join(' ')}
      ref={paletteAnchorRef}
    >
      {/* Primary palette stays fixed; advanced is absolutely positioned so it never shifts this box */}
      <ColorPalettePanel
        recent={recent}
        onSelect={handlePaletteSelect}
        onDefault={handleDefault}
        onMore={() => setShowAdvanced(true)}
      />
      {showAdvanced ? (
        <div
          data-testid="advanced-panel-anchor"
          className={
            advancedSide === 'right'
              ? 'absolute bottom-0 left-full z-[1]'
              : 'absolute bottom-0 right-full z-[1]'
          }
        >
          <ColorAdvancedPanel value={value} onChange={handleAdvancedChange} />
        </div>
      ) : null}
    </div>
  );
}
