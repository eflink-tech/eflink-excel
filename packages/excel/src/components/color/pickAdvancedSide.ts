/** Advanced panel width (matches ColorAdvancedPanel `w-[280px]`) */
export const ADVANCED_PANEL_WIDTH = 280;
export const PANEL_GAP = 8;

export type PanelSide = 'left' | 'right';

/**
 * Decide which side of the (stationary) palette to place the advanced panel.
 * Prefer right when it fits; otherwise left when that fits; else the roomier side.
 * The palette itself must not move.
 */
export function pickAdvancedSide(args: {
  paletteLeft: number;
  paletteRight: number;
  viewportWidth: number;
  advancedWidth?: number;
  gap?: number;
}): PanelSide {
  const width = args.advancedWidth ?? ADVANCED_PANEL_WIDTH;
  const gap = args.gap ?? PANEL_GAP;
  const spaceRight = args.viewportWidth - args.paletteRight - gap;
  const spaceLeft = args.paletteLeft - gap;

  if (spaceRight >= width) return 'right';
  if (spaceLeft >= width) return 'left';
  return spaceLeft > spaceRight ? 'left' : 'right';
}
