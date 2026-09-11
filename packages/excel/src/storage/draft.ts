// 本地草稿：手动保存策略下的兜底。
// 内容变化后防抖写入 localStorage（eflink:draft:excel:<docId>），云端保存成功后清除；
// 云端加载失败（网络异常）时用它回退恢复表格。读写全部静默失败，不影响编辑主流程。
import type { WorkbookSnapshot } from '../types/spreadsheet';

const DRAFT_PREFIX = 'eflink:draft:excel:';

function draftKey(docId: string): string {
  return DRAFT_PREFIX + docId;
}

/** 防抖到期写入本地草稿（不清 dirty、不触碰云端）；隐私模式/容量超限等场景静默失败 */
export function writeDraft(docId: string, snapshot: WorkbookSnapshot): void {
  try {
    localStorage.setItem(draftKey(docId), JSON.stringify(snapshot));
  } catch {
    // 草稿只是兜底，写不进去不打断编辑
  }
}

/** 读取本地草稿；无草稿或解析失败返回 null */
export function readDraft(docId: string): WorkbookSnapshot | null {
  try {
    const raw = localStorage.getItem(draftKey(docId));
    if (!raw) return null;
    return JSON.parse(raw) as WorkbookSnapshot;
  } catch {
    return null;
  }
}

/** 清除本地草稿（云端保存成功 / 宿主通过 bridge 丢弃草稿时调用） */
export function clearDraft(docId: string): void {
  try {
    localStorage.removeItem(draftKey(docId));
  } catch {
    // 静默失败
  }
}
