// 底部状态栏：文档名（点击改名）+ 保存状态。始终渲染（不随 branding 隐藏）——名字与保存状态是编辑器自身信息。
import { useEffect, useRef, useState } from 'react';
import { Pencil } from 'lucide-react';
import { useEditorStore } from '../store/editorStore';
import { useDocumentsStore } from '../store/documentsStore';
import { useUiStore } from '../store/uiStore';

/** 保存状态指示：有未保存改动时红点提示，云端保存成功后显示已保存（自 TitleBar 迁入） */
function SaveStatus() {
  const dirty = useEditorStore((s) => s.dirty);
  if (dirty) {
    return (
      <span className="flex shrink-0 items-center gap-1 text-[#e02e2e]" title="有未保存的修改，按 ⌘S / Ctrl+S 保存">
        <span className="text-[10px] leading-none">●</span>
        未保存
      </span>
    );
  }
  return (
    <span className="flex shrink-0 items-center gap-1">
      <span className="text-[10px] leading-none">✓</span>
      已保存
    </span>
  );
}

export function BottomBar() {
  const docId = useEditorStore((s) => s.docId);
  const editorTitle = useEditorStore((s) => s.title);
  // 文档名以 documentsStore 为准（renameDoc 后 refresh 生效）；独立打开编辑器等 docs 未加载的场景回退编辑器标题
  const docsTitle = useDocumentsStore((s) => s.docs.find((d) => d.id === docId)?.title);
  const title = docsTitle ?? editorTitle ?? '未命名表格';

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);
  // 单次编辑守卫：Enter 提交后输入框卸载触发的 blur、Esc 取消后的 blur 不再重复提交
  const doneRef = useRef(false);

  useEffect(() => {
    if (editing) {
      inputRef.current?.focus();
      inputRef.current?.select();
    }
  }, [editing]);

  const startEdit = () => {
    doneRef.current = false;
    setDraft(title);
    setEditing(true);
  };

  // 提交改名：空值不提交；云端 rename 成功后同步编辑器标题（保存链路用的是 editorStore.title，不同步会被下次保存覆盖回旧名）
  const commit = () => {
    if (doneRef.current) return;
    doneRef.current = true;
    setEditing(false);
    const next = draft.trim();
    if (!next || !docId || next === title) return;
    void useDocumentsStore
      .getState()
      .renameDoc(docId, next)
      .then(() => useEditorStore.getState().setTitle(next))
      .catch((err) => {
        console.error('重命名失败', err);
        useUiStore.getState().showToast('重命名失败');
      });
  };

  const cancel = () => {
    if (doneRef.current) return;
    doneRef.current = true;
    setEditing(false);
  };

  return (
    <div className="flex h-7 shrink-0 select-none items-center justify-between border-t border-[#e0e0e0] bg-[#f8f8f8] px-3 text-[11px] text-[#888]">
      {/* 左侧：文档名（点击改名）+ 保存状态 */}
      <div className="flex min-w-0 items-center gap-3">
        {editing ? (
          <input
            ref={inputRef}
            value={draft}
            maxLength={100}
            onChange={(e) => setDraft(e.target.value)}
            onBlur={commit}
            onKeyDown={(e) => {
              if (e.key === 'Enter') commit();
              else if (e.key === 'Escape') cancel();
            }}
            className="h-4 w-[200px] rounded border border-[#2b7fff] bg-white px-1 text-[11px] leading-none text-[#333] outline-none"
          />
        ) : (
          <button
            type="button"
            onClick={startEdit}
            title="点击重命名"
            className="flex max-w-[200px] items-center gap-1 rounded px-1 py-0.5 text-[#333] transition-colors hover:bg-[#e8e8e8]"
          >
            <Pencil size={11} className="shrink-0 text-[#999]" />
            <span className="truncate">{title}</span>
          </button>
        )}
        <SaveStatus />
      </div>
    </div>
  );
}
