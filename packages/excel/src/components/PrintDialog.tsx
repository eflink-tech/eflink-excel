// 打印 / PDF 弹窗：预览当前工作表已用区域，可切换纸张方向；
// 「打印」与「另存为 PDF」都走系统打印对话框（浏览器原生支持导出 PDF）。
import { useEffect, useMemo, useState } from 'react';
import { Printer } from 'lucide-react';
import { useI18n } from '@eflink-tech/office-kit';
import { currentSnapshot } from '../core/saveService';
import { getActiveSheetImageData } from '../core/univer/controller';
import { buildPrintHtml, printHtml, type PrintOrientation } from '../core/print/printSheet';
import { useEditorStore } from '../store/editorStore';

interface PrintDialogProps {
  onClose: () => void;
}

export function PrintDialog({ onClose }: PrintDialogProps) {
  const { t } = useI18n();
  const title = useEditorStore((s) => s.title);
  const [orientation, setOrientation] = useState<PrintOrientation>('portrait');

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onClose]);

  const data = useMemo(() => {
    const snapshot = currentSnapshot();
    return snapshot ? getActiveSheetImageData(snapshot) : null;
  }, []);

  const empty = useMemo(
    () => !data || data.displayValues.every((row) => row.every((value) => value === '')),
    [data],
  );

  const html = useMemo(() => {
    if (!data || empty) return '';
    return buildPrintHtml(data, {
      orientation,
      title: title || t('excel.untitled'),
      sheetName: data.sheet.name,
    });
  }, [data, empty, orientation, title, t]);

  const doPrint = () => {
    if (html) printHtml(html);
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
        aria-label={t('excel.print.title')}
        className="flex max-h-[90vh] flex-col overflow-hidden rounded-xl bg-white shadow-2xl"
        style={{ width: 920, maxWidth: 'calc(100vw - 32px)' }}
      >
        <div className="flex items-center justify-between border-b border-slate-100 px-5 py-3">
          <span className="flex items-center gap-2 text-sm font-semibold text-slate-700">
            <Printer size={16} />
            {t('excel.print.title')}
          </span>
          <button
            type="button"
            className="rounded p-1 text-slate-400 hover:bg-slate-100"
            aria-label={t('action.close')}
            title={t('action.close')}
            onClick={onClose}
          >
            ✕
          </button>
        </div>

        <div className="min-h-0 flex-1 overflow-y-auto p-5">
          <div className="mb-3 flex flex-wrap items-center gap-x-5 gap-y-2 text-xs text-slate-600">
            <span className="font-medium text-slate-700">{t('excel.print.orientation')}</span>
            {(['portrait', 'landscape'] as const).map((value) => (
              <label key={value} className="flex cursor-pointer items-center gap-1.5">
                <input
                  type="radio"
                  name="print-orientation"
                  value={value}
                  checked={orientation === value}
                  onChange={() => setOrientation(value)}
                />
                {value === 'portrait' ? t('excel.print.portrait') : t('excel.print.landscape')}
              </label>
            ))}
            <span className="text-slate-400">{t('excel.print.scope')}</span>
          </div>

          <div className="mb-2 text-xs font-medium text-slate-500">{t('excel.print.preview')}</div>
          {empty ? (
            <div className="flex h-64 items-center justify-center rounded-lg border border-dashed border-slate-300 bg-slate-50 text-sm text-slate-400">
              {t('excel.print.empty')}
            </div>
          ) : (
            <iframe
              title={t('excel.print.preview')}
              srcDoc={html}
              className="h-[52vh] w-full rounded-lg border border-slate-200 bg-white"
            />
          )}
          <div className="mt-3 rounded-lg bg-blue-50 px-3 py-2 text-xs leading-5 text-blue-700">
            {t('excel.print.hint')}
          </div>
        </div>

        <div className="flex justify-end gap-2 border-t border-slate-100 px-5 py-3">
          <button
            type="button"
            className="rounded-lg border border-slate-300 px-4 py-2 text-sm text-slate-700 hover:bg-slate-50"
            onClick={onClose}
          >
            {t('excel.confirm.cancel')}
          </button>
          <button
            type="button"
            className="rounded-lg border border-[#2b7fff] px-4 py-2 text-sm text-[#2b7fff] hover:bg-blue-50 disabled:cursor-not-allowed disabled:opacity-50"
            disabled={empty}
            onClick={doPrint}
          >
            {t('excel.print.pdf')}
          </button>
          <button
            type="button"
            className="rounded-lg bg-[#2b7fff] px-4 py-2 text-sm text-white hover:bg-[#1f6ff2] disabled:cursor-not-allowed disabled:opacity-50"
            disabled={empty}
            onClick={doPrint}
          >
            {t('excel.print.print')}
          </button>
        </div>
      </div>
    </div>
  );
}
