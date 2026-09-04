import { X } from 'lucide-react';
import { useEffect } from 'react';
import { createPortal } from 'react-dom';

interface ShortcutEntry {
  label: string;
  keys: string;
}

interface ShortcutGroup {
  title: string;
  entries: ShortcutEntry[];
}

/** 左 9 条 / 右 9 条，按条目数均衡分列 */
const COLUMNS: ShortcutGroup[][] = [
  [
    {
      title: '文件',
      entries: [{ label: '保存到本地数据库', keys: '⌘S' }],
    },
    {
      title: '编辑',
      entries: [
        { label: '撤销', keys: '⌘Z' },
        { label: '重做', keys: '⌘⇧Z / ⌘Y' },
        { label: '复制', keys: '⌘C' },
        { label: '剪切', keys: '⌘X' },
        { label: '粘贴', keys: '⌘V' },
        { label: '加粗', keys: '⌘B' },
        { label: '斜体', keys: '⌘I' },
        { label: '下划线', keys: '⌘U' },
      ],
    },
  ],
  [
    {
      title: '单元格',
      entries: [
        { label: '编辑单元格', keys: '双击 / Enter' },
        { label: '结束编辑', keys: 'Enter / Tab' },
        { label: '取消编辑', keys: 'Esc' },
        { label: '单元格内换行', keys: 'Alt + Enter' },
        { label: '删除并开始编辑', keys: 'Delete' },
      ],
    },
    {
      title: '导航与选区',
      entries: [
        { label: '全选', keys: '⌘A' },
        { label: '移动选区', keys: '↑ ↓ ← →' },
        { label: '扩展选区', keys: '⇧ + 方向键' },
        { label: '选择下一行首/行尾', keys: 'Home / End' },
      ],
    },
  ],
];

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
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [open, onClose]);

  if (!open) return null;

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
            快捷键
          </h2>
          <button
            type="button"
            aria-label="关闭"
            className="flex h-8 w-8 items-center justify-center rounded-md text-slate-500 hover:bg-slate-100"
            onClick={onClose}
          >
            <X size={18} />
          </button>
        </div>

        <div className="overflow-y-auto px-5 py-4">
          <p className="mb-4 text-center text-[12px] text-slate-400">
            Mac 以 ⌘ 表示 Command，Windows 对应 Ctrl
          </p>

          <div className="shortcuts-dialog__columns">
            {COLUMNS.map((column, colIdx) => (
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
