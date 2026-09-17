// 顶部菜单可用的文件操作（标题栏、AppMenu 共用）
import { currentSnapshot, saveNow } from './saveService';
import { exportEfexcel, importEfexcel } from './efexcel';
import { exportPng, exportXlsx, importXlsx, trimExt } from './xlsx/fileIO';
import { pickFile } from './filePicker';
import { getDefaultStorage } from '../storage/registry';
import { clearDraft } from '../storage/draft';
import { useDocumentsStore } from '../store/documentsStore';
import { useEditorStore } from '../store/editorStore';
import { useUiStore } from '../store/uiStore';

export async function newDocAction(): Promise<void> {
  // 二次确认：提醒用户先备份（导出 efexcel）再新建
  const ok = await useUiStore.getState().requestConfirm({
    title: '新建文档',
    message: '新建文档将清空当前内容，请确保已经下载备份文档！',
  });
  if (!ok) return;
  const doc = await useDocumentsStore.getState().createDoc();
  useEditorStore.getState().openDoc({ id: doc.id, title: doc.title });
}

export async function saveAction(): Promise<void> {
  try {
    await saveNow();
    useUiStore.getState().showToast('已保存');
  } catch (err) {
    console.error('保存失败', err);
    useUiStore.getState().showToast('保存失败');
  }
}

export async function exportEfexcelAction(): Promise<void> {
  const { docId } = useEditorStore.getState();
  if (!docId || !currentSnapshot()) return;
  try {
    await saveNow();
    const doc = await getDefaultStorage().load(docId);
    if (!doc) throw new Error(`文档不存在: ${docId}`);
    exportEfexcel(doc);
    useUiStore.getState().showToast('已导出 efexcel');
  } catch (err) {
    console.error('导出失败', err);
    useUiStore.getState().showToast('导出失败');
  }
}

/** 导出图片是纯读操作：直接截当前 DOM，不保存、不改变 dirty 状态 */
export async function exportPngAction(): Promise<void> {
  const el = document.getElementById('univer-container');
  if (!el) return;
  const { title } = useEditorStore.getState();
  try {
    await exportPng(el, title || '未命名表格');
    useUiStore.getState().showToast('已导出图片');
  } catch (err) {
    console.error('导出图片失败', err);
    useUiStore.getState().showToast('导出失败');
  }
}

export async function importFileAction(file: File): Promise<void> {
  try {
    const doc = await importEfexcel(file);
    await useDocumentsStore.getState().addDoc(doc);
    useEditorStore.getState().openDoc({ id: doc.id, title: doc.title });
    useUiStore.getState().showToast('导入成功');
  } catch (err) {
    console.error('导入失败', err);
    useUiStore.getState().showToast('导入失败：无法解析该文件');
  }
}

/** 导出 Excel (.xlsx)：先保存，再取最新快照转换下载 */
export async function exportXlsxAction(): Promise<void> {
  const { docId } = useEditorStore.getState();
  if (!docId || !currentSnapshot()) return;
  try {
    await saveNow();
    const doc = await getDefaultStorage().load(docId);
    if (!doc) throw new Error(`文档不存在: ${docId}`);
    await exportXlsx(doc.snapshot, doc.title);
    useUiStore.getState().showToast('已导出 xlsx');
  } catch (err) {
    console.error('导出 xlsx 失败', err);
    useUiStore.getState().showToast('导出失败');
  }
}

/** 导入 Excel (.xlsx)：解析成功后替换当前文档内容（docId 不变），并触发编辑器重载 */
export async function importXlsxAction(file: File): Promise<void> {
  const { docId, title } = useEditorStore.getState();
  if (!docId) return;
  const ok = await useUiStore.getState().requestConfirm({
    title: '导入 Excel',
    message: `导入将替换当前文档「${title}」的内容，确定继续吗？`,
  });
  if (!ok) return;
  // 确认期间文档可能已被切换/关闭，重新快照 docId（TOCTOU）
  const confirmedDocId = useEditorStore.getState().docId;
  if (!confirmedDocId) return;
  try {
    // 先保存：冲掉 dirty 与自动保存防抖，避免导入后 ⌘S/自动保存用旧编辑器快照覆盖导入结果
    await saveNow();
  } catch (err) {
    console.error('导入前保存失败', err);
    useUiStore.getState().showToast('导入失败：保存文档失败');
    return;
  }
  try {
    // 先解析：失败在落盘前抛出，当前文档不受影响
    const snapshot = await importXlsx(file);
    try {
      await getDefaultStorage().updateContent(confirmedDocId, trimExt(file.name), snapshot);
    } catch (err) {
      console.error('导入保存文档失败', err);
      useUiStore.getState().showToast('导入失败：保存文档失败');
      return;
    }
    clearDraft(confirmedDocId); // 防重载读到旧草稿
    void useDocumentsStore.getState().refresh().catch(() => {}); // 防列表刷新失败的 unhandledrejection
    useEditorStore.getState().bumpReload();
    useUiStore.getState().showToast('导入成功');
  } catch (err) {
    console.error('导入 Excel 失败', err);
    useUiStore.getState().showToast('导入失败：无法解析该文件');
  }
}

/** 菜单入口：弹文件选择框（限 xlsx/xlsm），选中后走导入动作 */
export async function importXlsxMenuAction(): Promise<void> {
  const file = await pickFile('.xlsx,.xlsm');
  if (!file) return;
  await importXlsxAction(file);
}
