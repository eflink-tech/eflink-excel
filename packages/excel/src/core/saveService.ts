// 保存服务：绑定当前编辑文档，统一“手动保存 + 本地草稿”入口（TitleBar、快捷键、分享前共用）。
// 策略：内容变化 → 置 dirty + 防抖 1.5s 写本地草稿（不清 dirty、不动云端）；
// ⌘S / 分享前 → 云端保存，成功后才清 dirty 并删除本地草稿。
import { createAutosaver } from './autosave';
import { getDefaultStorage } from '../storage/registry';
import { clearDraft, writeDraft } from '../storage/draft';
import { useEditorStore } from '../store/editorStore';
import type { WorkbookSnapshot } from '../types/spreadsheet';

let autosaver = createAutosaver(writeDraftSnapshot, 1500);
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

/** Univer onCommandExecuted → 内容变化入口：置脏 + 防抖写本地草稿 */
export function notifyChanged(): void {
  if (!bound) return;
  useEditorStore.getState().markDirty();
  autosaver.schedule();
}

/** 防抖到期：仅写本地草稿（静默失败），不调用云端保存、不清 dirty */
function writeDraftSnapshot(): void {
  if (!bound) return;
  const snapshot = bound.getSnapshot();
  if (!snapshot) return;
  writeDraft(bound.docId, snapshot);
}

/** ⌘S / 分享前入口：云端保存；成功后 dirty=false 并删除本地草稿 */
export async function saveNow(): Promise<void> {
  if (!bound) return;
  autosaver.cancel();
  // 云端保存前先同步写一次草稿兜底：保存失败（断网等）时草稿保持最新
  if (useEditorStore.getState().dirty) writeDraftSnapshot();
  await doSave();
}

async function doSave(): Promise<void> {
  if (!bound) return;
  const snapshot = bound.getSnapshot();
  if (!snapshot) return;
  const editor = useEditorStore.getState();
  editor.setSaving(true);
  try {
    await getDefaultStorage().updateContent(bound.docId, editor.title, snapshot);
    clearDraft(bound.docId);
    useEditorStore.getState().markSaved();
  } catch (err) {
    useEditorStore.getState().setSaving(false);
    throw err;
  }
}
