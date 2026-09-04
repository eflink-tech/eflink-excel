// 顶部菜单可用的文件操作（标题栏、AppMenu 共用）
import { currentSnapshot, saveNow } from './saveService';
import { exportEfexcel, importEfexcel } from './efexcel';
import { exportPng } from './xlsx/fileIO';
import { getDefaultStorage } from '../storage/registry';
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

export async function exportPngAction(): Promise<void> {
  const el = document.getElementById('univer-container');
  if (!el) return;
  const { title } = useEditorStore.getState();
  try {
    await saveNow();
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
