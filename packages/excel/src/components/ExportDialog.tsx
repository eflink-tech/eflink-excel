// 导出弹窗：与 PPTX/Word 相同交互——一个「导出」入口，弹窗内选择文件类型
// - 导出期间按钮显示 spinner +「导出中」，并禁用其余选项
// - 图片需整表重绘，耗时较长，先让 loading 完成一次绘制再开始导出
import { useEffect, useState } from 'react';
import { flushSync } from 'react-dom';
import { Loader2 } from 'lucide-react';
import { useI18n } from '@eflink-tech/office-kit';
import { exportCsvAction, exportEfexcelAction, exportPngAction, exportXlsxAction } from '../core/fileActions';

interface ExportDialogProps {
  onClose: () => void;
  /** 从导出弹窗进入打印 / PDF 预览 */
  onOpenPrint?: () => void;
}

type ExportKey = 'xlsx' | 'csv' | 'pdf' | 'image' | 'data';

// 双 rAF 等 loading 真正绘制一帧后再开始重导出，否则同步绘制会阻塞首帧
function nextPaint(): Promise<void> {
  return new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
  });
}

export function ExportDialog({ onClose, onOpenPrint }: ExportDialogProps) {
  const { t } = useI18n();
  const [busy, setBusy] = useState<ExportKey | null>(null);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !busy) onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [busy, onClose]);

  const run = async (key: ExportKey, fn: () => Promise<void>) => {
    if (busy) return;
    flushSync(() => setBusy(key));
    await nextPaint();
    try {
      await fn();
    } finally {
      setBusy(null);
    }
  };

  const options: Array<{ key: ExportKey; title: string; desc: string; run: () => Promise<void> }> = [
    { key: 'xlsx', title: t('excel.export.xlsx'), desc: t('excel.export.xlsxDesc'), run: exportXlsxAction },
    { key: 'csv', title: t('excel.export.csv'), desc: t('excel.export.csvDesc'), run: exportCsvAction },
    {
      key: 'pdf',
      title: t('excel.export.pdf'),
      desc: t('excel.export.pdfDesc'),
      run: async () => {
        onOpenPrint?.();
      },
    },
    { key: 'image', title: t('excel.export.image'), desc: t('excel.export.imageDesc'), run: exportPngAction },
    { key: 'data', title: t('excel.export.data'), desc: t('excel.export.dataDesc'), run: exportEfexcelAction },
  ];

  return (
    <div
      className="fixed inset-0 z-[10000] flex items-center justify-center bg-black/45"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget && !busy) onClose();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-label={t('excel.export.title')}
        className="max-h-[85vh] overflow-hidden rounded-xl bg-white shadow-2xl"
        style={{ width: 520 }}
      >
        <div className="flex items-center justify-between border-b border-slate-100 px-5 py-3">
          <span className="text-sm font-semibold text-slate-700">{t('excel.export.title')}</span>
          <button
            type="button"
            className="rounded p-1 text-slate-400 hover:bg-slate-100 disabled:opacity-50"
            aria-label={t('action.close')}
            title={t('action.close')}
            disabled={!!busy}
            onClick={onClose}
          >
            ✕
          </button>
        </div>
        <div className="max-h-[75vh] overflow-y-auto p-5">
          <div className="space-y-2.5">
            {options.map((option) => (
              <button
                key={option.key}
                type="button"
                className="w-full rounded-lg border border-gray-200 px-4 py-3 text-left hover:border-[#2b7fff] disabled:cursor-not-allowed disabled:opacity-50"
                disabled={!!busy}
                onClick={() => void run(option.key, option.run)}
              >
                <div className="flex items-center gap-2 text-sm font-medium text-gray-700">
                  {busy === option.key ? (
                    <span className="inline-flex items-center gap-1.5">
                      <Loader2 size={14} className="animate-spin" />
                      {t('excel.export.busy')}
                    </span>
                  ) : (
                    option.title
                  )}
                </div>
                <div className="mt-0.5 text-xs text-gray-400">{option.desc}</div>
              </button>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}
