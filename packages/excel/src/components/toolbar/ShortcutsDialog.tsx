import { X } from 'lucide-react';
import { useEffect } from 'react';
import { createPortal } from 'react-dom';
import { useI18n } from '@eflink-tech/office-kit';
// 注册表格专属文案，保证组件脱离 I18nProvider 单独渲染时也有中文兜底
import '../../i18n';

interface ShortcutEntry {
  label: string;
  keys: string;
}

interface ShortcutGroup {
  title: string;
  entries: ShortcutEntry[];
}

function buildColumns(t: (key: string) => string): ShortcutGroup[][] {
  return [
    [
      {
        title: t('excel.shortcuts.file'),
        entries: [{ label: t('excel.shortcuts.saveLocal'), keys: '⌘S' }],
      },
      {
        title: t('excel.shortcuts.edit'),
        entries: [
          { label: t('action.undo'), keys: '⌘Z' },
          { label: t('action.redo'), keys: '⌘⇧Z / ⌘Y' },
          { label: t('action.copy'), keys: '⌘C' },
          { label: t('action.cut'), keys: '⌘X' },
          { label: t('action.paste'), keys: '⌘V' },
          { label: t('excel.toolbar.bold'), keys: '⌘B' },
          { label: t('excel.toolbar.italic'), keys: '⌘I' },
          { label: t('excel.toolbar.underline'), keys: '⌘U' },
        ],
      },
    ],
    [
      {
        title: t('excel.shortcuts.cell'),
        entries: [
          { label: t('excel.shortcuts.editCell'), keys: `${t('excel.shortcuts.doubleClick')} / Enter` },
          { label: t('excel.shortcuts.finishEdit'), keys: 'Enter / Tab' },
          { label: t('excel.shortcuts.cancelEdit'), keys: 'Esc' },
          { label: t('excel.shortcuts.newline'), keys: 'Alt + Enter' },
          { label: t('excel.shortcuts.deleteEdit'), keys: 'Delete' },
        ],
      },
      {
        title: t('excel.shortcuts.navigation'),
        entries: [
          { label: t('action.selectAll'), keys: '⌘A' },
          { label: t('excel.shortcuts.moveSelection'), keys: '↑ ↓ ← →' },
          { label: t('excel.shortcuts.extendSelection'), keys: `⇧ + ${t('excel.shortcuts.arrowKeys')}` },
          { label: t('excel.shortcuts.rowStartEnd'), keys: 'Home / End' },
        ],
      },
    ],
  ];
}

function ShortcutGroupBlock({ group }: { group: ShortcutGroup }) {
  return (
    <section className="shortcuts-dialog__group">
      <h3 className="shortcuts-dialog__group-title">{group.title}</h3>
      <div className="shortcuts-dialog__list">
        {group.entries.map((entry, i) => (
          <div
            key={entry.label}
            className={`shortcuts-dialog__row${i > 0 ? ' shortcuts-dialog__row--border' : ''}`}
          >
            <span className="shortcuts-dialog__label">{entry.label}</span>
            <kbd className="shortcuts-dialog__keys">{entry.keys}</kbd>
          </div>
        ))}
      </div>
    </section>
  );
}

/** 快捷键总览弹窗（参照思维导图的快捷键弹窗样式） */
export function ShortcutsDialog({ open, onClose }: { open: boolean; onClose: () => void }) {
  const { t } = useI18n();
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [open, onClose]);

  if (!open) return null;

  const columns = buildColumns(t);

  return createPortal(
    <div
      className="shortcuts-dialog-overlay fixed inset-0 z-[100] flex items-center justify-center bg-black/40 p-4"
      onClick={onClose}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="shortcuts-dialog-title"
        className="shortcuts-dialog flex max-h-[80vh] w-[640px] max-w-full flex-col overflow-hidden rounded-xl bg-white shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex shrink-0 items-center justify-between border-b border-slate-100 px-5 py-3.5">
          <h2 id="shortcuts-dialog-title" className="text-[16px] font-semibold text-slate-900">
            {t('excel.shortcuts.title')}
          </h2>
          <button
            type="button"
            aria-label={t('action.close')}
            className="flex h-8 w-8 items-center justify-center rounded-md text-slate-500 hover:bg-slate-100"
            onClick={onClose}
          >
            <X size={18} />
          </button>
        </div>

        <div className="overflow-y-auto px-5 py-4">
          <p className="mb-4 text-center text-[12px] text-slate-400">
            {t('excel.shortcuts.hint')}
          </p>

          <div className="shortcuts-dialog__columns">
            {columns.map((column, colIdx) => (
              <div key={colIdx} className="shortcuts-dialog__column">
                {column.map((group) => (
                  <ShortcutGroupBlock key={group.title} group={group} />
                ))}
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>,
    document.body,
  );
}
