// 导入弹窗：与 PPTX/Word 相同交互——选择文件 → 预览 → 确认导入
// - .xlsx / .xlsm：先渲染工作表预览，确认后替换当前文档内容
// - .json（.efx.json / .efexcel）：工程文件原样还原，直接导入为新文档
import { useEffect, useState, type ReactNode } from 'react';
import { useI18n } from '@eflink-tech/office-kit';
import { importFileAction, importXlsxAction } from '../core/fileActions';
import { parseXlsxSnapshot } from '../core/xlsx/fileIO';
import { useEditorStore } from '../store/editorStore';
import { XlsxPreview } from './XlsxPreview';
import type { WorkbookSnapshot } from '../types/spreadsheet';

interface ImportDialogProps {
  onClose: () => void;
}

interface XlsxPreviewState {
  file: File;
  snapshot: WorkbookSnapshot;
}

type Busy = null | 'preview' | 'import';

export function ImportDialog({ onClose }: ImportDialogProps) {
  const { t } = useI18n();
  const title = useEditorStore((s) => s.title);
  const [preview, setPreview] = useState<XlsxPreviewState | null>(null);
  const [busy, setBusy] = useState<Busy>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !busy) onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [busy, onClose]);

  const handleFile = async (file: File) => {
    setError(null);
    const lower = file.name.toLowerCase();
    if (lower.endsWith('.json') || lower.endsWith('.efexcel')) {
      setBusy('import');
      const ok = await importFileAction(file);
      if (ok) {
        onClose();
      } else {
        setBusy(null);
      }
      return;
    }
    if (!lower.endsWith('.xlsx') && !lower.endsWith('.xlsm')) {
      setError(t('excel.import.chooseFile'));
      return;
    }
    setBusy('preview');
    try {
      const snapshot = await parseXlsxSnapshot(file);
      setPreview({ file, snapshot });
    } catch (err) {
      console.error('解析 Excel 预览失败', err);
      setError(t('excel.import.previewFailed'));
    } finally {
      setBusy(null);
    }
  };

  const handleConfirm = async () => {
    if (!preview || busy) return;
    setBusy('import');
    setError(null);
    const ok = await importXlsxAction(preview.file);
    if (ok) {
      onClose();
    } else {
      setBusy(null);
    }
  };

  const shell = (heading: string, width: number, body: ReactNode) => (
    <div
      className="fixed inset-0 z-[10000] flex items-center justify-center bg-black/45"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget && !busy) onClose();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-label={heading}
        className="max-h-[85vh] overflow-hidden rounded-xl bg-white shadow-2xl"
        style={{ width }}
      >
        <div className="flex items-center justify-between border-b border-slate-100 px-5 py-3">
          <span className="text-sm font-semibold text-slate-700">{heading}</span>
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
        <div className="max-h-[75vh] overflow-y-auto p-5">{body}</div>
      </div>
    </div>
  );

  if (preview) {
    return shell(
      t('excel.import.previewTitle'),
      860,
      <>
        <XlsxPreview snapshot={preview.snapshot} />
        <div className="mt-3 rounded-lg bg-amber-50 px-3 py-2 text-xs leading-5 text-amber-700">
          {t('excel.import.replaceNote', { title: title || t('excel.untitled') })}
        </div>
        {error && <div className="mt-2 text-xs text-red-500">{error}</div>}
        <div className="mt-4 flex justify-end gap-2">
          <button
            type="button"
            className="rounded-lg border border-slate-300 px-4 py-2 text-sm text-slate-700 hover:bg-slate-50 disabled:opacity-50"
            disabled={!!busy}
            onClick={() => {
              setPreview(null);
              setError(null);
            }}
          >
            {t('excel.import.back')}
          </button>
          <button
            type="button"
            className="rounded-lg bg-[#2b7fff] px-4 py-2 text-sm text-white hover:bg-[#1f6ff2] disabled:opacity-50"
            disabled={!!busy}
            onClick={() => void handleConfirm()}
          >
            {t('excel.import.confirm')}
          </button>
        </div>
      </>,
    );
  }

  return shell(
    t('excel.import.title'),
    520,
    <>
      <label className="block cursor-pointer rounded-xl border-2 border-dashed border-slate-300 px-6 py-10 text-center hover:border-[#2b7fff]">
        <input
          type="file"
          accept=".xlsx,.xlsm,.json"
          className="hidden disabled:opacity-50"
          disabled={!!busy}
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) void handleFile(file);
            e.target.value = '';
          }}
        />
        <div className="text-sm text-slate-600">
          {busy === 'preview' ? t('excel.import.previewing') : t('excel.import.choose')}
        </div>
        <div className="mt-1 text-xs text-slate-400">{t('excel.import.support')}</div>
        <div className="mt-2 text-[11px] text-slate-400">{t('excel.import.previewNote')}</div>
        <div className="mt-2 text-[11px] text-slate-400">{t('excel.import.limitNote')}</div>
      </label>
      {error && <div className="mt-3 text-xs text-red-500">{error}</div>}
    </>,
  );
}
