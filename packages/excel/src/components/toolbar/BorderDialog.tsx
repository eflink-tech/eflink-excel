// 边框设置弹窗：替换原「更多边框设置」占位，按勾选的位置批量应用线型与颜色
import { useEffect, useState } from 'react';
import { useI18n } from '@eflink-tech/office-kit';
import type { BorderStyleKey, BorderTypeKey } from '../../core/univer/commands';

export interface BorderDialogProps {
  open: boolean;
  initialStyle: BorderStyleKey;
  initialColor: string;
  onClose: () => void;
  onApply: (edges: BorderTypeKey[], style: BorderStyleKey, color: string) => void;
  onClear: () => void;
}

const EDGES: { key: BorderTypeKey; labelKey: string }[] = [
  { key: 'top', labelKey: 'excel.border.preset.top' },
  { key: 'bottom', labelKey: 'excel.border.preset.bottom' },
  { key: 'left', labelKey: 'excel.border.preset.left' },
  { key: 'right', labelKey: 'excel.border.preset.right' },
  { key: 'horizontal', labelKey: 'excel.border.edge.horizontal' },
  { key: 'vertical', labelKey: 'excel.border.edge.vertical' },
  { key: 'tlbr', labelKey: 'excel.border.preset.tlbr' },
  { key: 'bltr', labelKey: 'excel.border.preset.bltr' },
];

const DEFAULT_EDGES: BorderTypeKey[] = ['top', 'bottom', 'left', 'right', 'horizontal', 'vertical'];

const STYLES: { key: BorderStyleKey; labelKey: string; width: number; dasharray?: string }[] = [
  { key: 'thin', labelKey: 'excel.border.style.thin', width: 1 },
  { key: 'medium', labelKey: 'excel.border.style.medium', width: 2.5 },
  { key: 'dashed', labelKey: 'excel.border.style.dashed', width: 1.5, dasharray: '4 3' },
  { key: 'dotted', labelKey: 'excel.border.style.dotted', width: 1.5, dasharray: '1.5 2.5' },
  { key: 'double', labelKey: 'excel.border.style.double', width: 3, dasharray: '3 1.5 0.8 1.5' },
];

const COLORS = ['#000000', '#595959', '#e33e33', '#e6821e', '#d8b100', '#41a62a', '#1e7ee6', '#7c31d6', '#9c1f5f', '#21a366'];

export function BorderDialog({ open, initialStyle, initialColor, onClose, onApply, onClear }: BorderDialogProps) {
  const { t } = useI18n();
  const [edges, setEdges] = useState<BorderTypeKey[]>(DEFAULT_EDGES);
  const [style, setStyle] = useState<BorderStyleKey>(initialStyle);
  const [color, setColor] = useState(initialColor);

  useEffect(() => {
    if (!open) return;
    setEdges(DEFAULT_EDGES);
    setStyle(initialStyle);
    setColor(initialColor);
  }, [open, initialStyle, initialColor]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open, onClose]);

  if (!open) return null;

  const toggleEdge = (key: BorderTypeKey) => {
    setEdges((prev) => (prev.includes(key) ? prev.filter((k) => k !== key) : [...prev, key]));
  };

  const apply = () => {
    if (!edges.length) return;
    onApply(edges, style, color);
    onClose();
  };

  return (
    <div
      className="fixed inset-0 z-[10000] flex items-center justify-center bg-black/45"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-label={t('excel.border.dialog.title')}
        className="w-[420px] rounded-xl bg-white p-5 shadow-2xl"
      >
        <div className="mb-3 flex items-center justify-between">
          <span className="text-sm font-semibold text-slate-700">{t('excel.border.dialog.title')}</span>
          <button type="button" className="rounded p-1 text-slate-400 hover:bg-slate-100" onClick={onClose} aria-label={t('action.close')}>
            ✕
          </button>
        </div>

        <div className="mb-1 text-xs text-slate-500">{t('excel.border.dialog.edges')}</div>
        <div className="mb-3 grid grid-cols-3 gap-1.5">
          {EDGES.map((edge) => {
            const checked = edges.includes(edge.key);
            return (
              <label
                key={edge.key}
                className={`flex cursor-pointer items-center gap-1.5 rounded border px-2 py-1 text-xs ${
                  checked ? 'border-[#2b7fff] bg-blue-50 text-[#1d6fe0]' : 'border-slate-200 text-slate-600'
                }`}
              >
                <input
                  type="checkbox"
                  className="h-3.5 w-3.5 accent-[#2b7fff]"
                  checked={checked}
                  onChange={() => toggleEdge(edge.key)}
                />
                <span className="truncate">{t(edge.labelKey)}</span>
              </label>
            );
          })}
        </div>

        <div className="mb-1 text-xs text-slate-500">{t('excel.border.dialog.style')}</div>
        <div className="mb-3 flex flex-wrap gap-1.5">
          {STYLES.map((s) => (
            <button
              key={s.key}
              type="button"
              className={`flex items-center gap-2 rounded border px-2 py-1 text-xs ${
                style === s.key ? 'border-[#2b7fff] bg-blue-50 text-[#1d6fe0]' : 'border-slate-200 text-slate-600'
              }`}
              onClick={() => setStyle(s.key)}
            >
              <svg width="30" height="8" viewBox="0 0 30 8">
                <path d="M1 4 H29" stroke={style === s.key ? '#1d4ed8' : '#475569'} strokeWidth={s.width} strokeDasharray={s.dasharray} />
              </svg>
              {t(s.labelKey)}
            </button>
          ))}
        </div>

        <div className="mb-1 text-xs text-slate-500">{t('excel.border.dialog.color')}</div>
        <div className="mb-4 flex flex-wrap gap-1.5">
          {COLORS.map((c) => (
            <button
              key={c}
              type="button"
              title={c}
              aria-label={c}
              className={`h-6 w-6 rounded border-2 ${color === c ? 'border-[#2b7fff]' : 'border-transparent'}`}
              style={{ background: c }}
              onClick={() => setColor(c)}
            />
          ))}
        </div>

        <div className="mb-4 rounded bg-slate-50 px-3 py-2 text-[11px] leading-5 text-slate-500">
          {t('excel.border.dialog.hint')}
        </div>

        <div className="flex justify-end gap-2">
          <button
            type="button"
            className="rounded border border-slate-200 px-3 py-1.5 text-sm text-slate-600 hover:bg-slate-50"
            onClick={() => {
              onClear();
              onClose();
            }}
          >
            {t('excel.border.dialog.clear')}
          </button>
          <button type="button" className="rounded border border-slate-200 px-3 py-1.5 text-sm text-slate-600 hover:bg-slate-50" onClick={onClose}>
            {t('excel.confirm.cancel')}
          </button>
          <button
            type="button"
            className="rounded bg-[#2b7fff] px-3 py-1.5 text-sm text-white hover:bg-[#1d6fe0] disabled:opacity-50"
            disabled={!edges.length}
            onClick={apply}
          >
            {t('excel.confirm.ok')}
          </button>
        </div>
      </div>
    </div>
  );
}
