// 自定义数字格式弹窗：替换原「更多格式」占位，输入 Excel 数字格式代码并应用到当前选区
import { useEffect, useState } from 'react';
import { useI18n } from '@eflink-tech/office-kit';

export interface NumberFormatDialogProps {
  open: boolean;
  /** 选区当前格式代码（空/General 表示常规） */
  value: string;
  onClose: () => void;
  onApply: (pattern: string) => void;
}

const TEMPLATES = [
  '0',
  '0.00',
  '#,##0',
  '#,##0.00',
  '0.00%',
  '"¥"#,##0.00',
  'yyyy-mm-dd',
  'yyyy"年"m"月"d"日"',
  'h:mm:ss',
  '@',
];

export function NumberFormatDialog({ open, value, onClose, onApply }: NumberFormatDialogProps) {
  const { t } = useI18n();
  const [draft, setDraft] = useState('');

  useEffect(() => {
    if (open) setDraft(value === 'General' ? '' : value);
  }, [open, value]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open, onClose]);

  if (!open) return null;

  const apply = () => {
    const pattern = draft.trim();
    if (!pattern) return;
    onApply(pattern);
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
        aria-label={t('excel.numfmt.dialogTitle')}
        className="w-[460px] rounded-xl bg-white p-5 shadow-2xl"
      >
        <div className="mb-3 flex items-center justify-between">
          <span className="text-sm font-semibold text-slate-700">{t('excel.numfmt.dialogTitle')}</span>
          <button type="button" className="rounded p-1 text-slate-400 hover:bg-slate-100" onClick={onClose} aria-label={t('action.close')}>
            ✕
          </button>
        </div>
        <label className="mb-1 block text-xs text-slate-500" htmlFor="numfmt-input">
          {t('excel.numfmt.code')}
        </label>
        <input
          id="numfmt-input"
          autoFocus
          value={draft}
          placeholder={t('excel.numfmt.placeholder')}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') apply();
          }}
          className="mb-3 w-full rounded border border-slate-300 px-2 py-1.5 text-sm text-slate-700 outline-none focus:border-[#2b7fff]"
        />
        <div className="mb-1 text-xs text-slate-500">{t('excel.numfmt.templates')}</div>
        <div className="mb-4 flex flex-wrap gap-1.5">
          {TEMPLATES.map((pattern) => (
            <button
              key={pattern}
              type="button"
              className="rounded border border-slate-200 bg-slate-50 px-2 py-1 text-xs text-slate-600 hover:border-[#2b7fff] hover:text-[#2b7fff]"
              onClick={() => setDraft(pattern)}
            >
              {pattern}
            </button>
          ))}
        </div>
        <div className="mb-4 rounded bg-slate-50 px-3 py-2 text-[11px] leading-5 text-slate-500">
          {t('excel.numfmt.hint')}
        </div>
        <div className="flex justify-end gap-2">
          <button type="button" className="rounded border border-slate-200 px-3 py-1.5 text-sm text-slate-600 hover:bg-slate-50" onClick={onClose}>
            {t('excel.confirm.cancel')}
          </button>
          <button
            type="button"
            className="rounded bg-[#2b7fff] px-3 py-1.5 text-sm text-white hover:bg-[#1d6fe0] disabled:opacity-50"
            disabled={!draft.trim()}
            onClick={apply}
          >
            {t('excel.confirm.ok')}
          </button>
        </div>
      </div>
    </div>
  );
}
