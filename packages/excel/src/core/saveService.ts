// 保存服务：绑定当前编辑文档，统一自动保存/手动保存入口（TitleBar、快捷键、卸载前共用）
import { createAutosaver } from './autosave';
import { getDefaultStorage } from '../storage/registry';
import { useEditorStore } from '../store/editorStore';
import type { WorkbookSnapshot } from '../types/spreadsheet';

let autosaver = createAutosaver(doSave);
let bound: { docId: string; getSnapshot: () => WorkbookSnapshot | null } | null = null;

export function bindEditor(docId: string, getSnapshot: () => WorkbookSnapshot | null): void {
  bound = { docId, getSnapshot };
}

export function unbindEditor(): void {
  autosaver.cancel();
  bound = null;
}

/** 当前文档的最新快照（导出等场景使用） */
export function currentSnapshot(): WorkbookSnapshot | null {
  return bound?.getSnapshot() ?? null;
}

/** Univer onCommandExecuted → 内容变化入口 */
export function notifyChanged(): void {
  if (!bound) return;
  useEditorStore.getState().markDirty();
  autosaver.schedule();
}

export async function saveNow(): Promise<void> {
  await autosaver.flush();
}

async function doSave(): Promise<void> {
  if (!bound) return;
  const snapshot = bound.getSnapshot();
  if (!snapshot) return;
  const editor = useEditorStore.getState();
  editor.setSaving(true);
  try {
    await getDefaultStorage().updateContent(bound.docId, editor.title, snapshot);
    useEditorStore.getState().markSaved();
  } catch (err) {
    useEditorStore.getState().setSaving(false);
    throw err;
  }
}
