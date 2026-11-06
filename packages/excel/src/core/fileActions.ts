// 顶部菜单可用的文件操作（标题栏、AppMenu 共用）
import { currentSnapshot, saveNow } from './saveService';
import { exportEfexcel, importEfexcel } from './efexcel';
import { exportXlsx, importXlsx, trimExt } from './xlsx/fileIO';
import { exportSheetPng } from './xlsx/sheetImage';
import { getActiveSheetImageData } from './univer/controller';
import { queueImportedCharts } from './xlsx/chartImport';
import { queueImportedAnnotations } from './xlsx/annotationImport';
import { getDefaultStorage } from '../storage/registry';
import { clearDraft } from '../storage/draft';
import { useDocumentsStore } from '../store/documentsStore';
import { useEditorStore } from '../store/editorStore';
import { useUiStore } from '../store/uiStore';
import { excelT } from '../i18n';

export async function newDocAction(): Promise<void> {
  // 二次确认：提醒用户先备份（导出 efexcel）再新建
  const ok = await useUiStore.getState().requestConfirm({
    title: excelT('excel.confirm.newDocTitle'),
    message: excelT('excel.confirm.newDocMessage'),
  });
  if (!ok) return;
  const doc = await useDocumentsStore.getState().createDoc();
  useEditorStore.getState().openDoc({ id: doc.id, title: doc.title });
}

export async function saveAction(): Promise<void> {
  try {
    await saveNow();
    useUiStore.getState().showToast(excelT('excel.toast.saved'));
  } catch (err) {
    console.error('保存失败', err);
    useUiStore.getState().showToast(excelT('excel.toast.saveFailed'));
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
    useUiStore.getState().showToast(excelT('excel.toast.exportedJson'));
  } catch (err) {
    console.error('导出失败', err);
    useUiStore.getState().showToast(excelT('excel.toast.exportFailed'));
  }
}

/** 导出图片：导出当前 sheet 的全部内容（按已用区域重建整表，而非可视区域截图） */
export async function exportPngAction(): Promise<void> {
  const { title } = useEditorStore.getState();
  const snapshot = currentSnapshot();
  const data = snapshot ? getActiveSheetImageData(snapshot) : null;
  if (!data) return;
  try {
    await exportSheetPng(data, title || excelT('excel.untitled'));
    useUiStore.getState().showToast(excelT('excel.toast.exportedPng'));
  } catch (err) {
    console.error('导出图片失败', err);
    useUiStore.getState().showToast(excelT('excel.toast.exportFailed'));
  }
}

export async function importFileAction(file: File): Promise<boolean> {
  try {
    const doc = await importEfexcel(file);
    await useDocumentsStore.getState().addDoc(doc);
    useEditorStore.getState().openDoc({ id: doc.id, title: doc.title });
    useUiStore.getState().showToast(excelT('excel.toast.imported'));
    return true;
  } catch (err) {
    console.error('导入失败', err);
    useUiStore.getState().showToast(excelT('excel.toast.importFailed'));
    return false;
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
    useUiStore.getState().showToast(excelT('excel.toast.exportedXlsx'));
  } catch (err) {
    console.error('导出 xlsx 失败', err);
    useUiStore.getState().showToast(excelT('excel.toast.exportFailed'));
  }
}

/** 超过该字节数的文档由宿主存本地浏览器，不写服务器数据库（与宿主 office 存储上限一致） */
const LOCAL_ONLY_LIMIT_BYTES = 500 * 1024;

/** 导入 Excel (.xlsx)：替换当前文档内容（docId 不变），并触发编辑器重载 */
export async function importXlsxAction(file: File): Promise<boolean> {
  const { docId } = useEditorStore.getState();
  if (!docId) return false;
  // 与 Word 导入一致：解析期间显示全屏 loading，结束（含失败）后关闭
  useUiStore.getState().setImporting(true);
  try {
    try {
      // 先保存：冲掉 dirty 与自动保存防抖，避免导入后 ⌘S/自动保存用旧编辑器快照覆盖导入结果
      await saveNow();
    } catch (err) {
      console.error('导入前保存失败', err);
      useUiStore.getState().showToast(excelT('excel.toast.importSaveFailed'));
      return false;
    }
    try {
      // 先解析：失败在落盘前抛出，当前文档不受影响（图表已渲染成图片随结果一并返回）
      const { snapshot, charts, annotations } = await importXlsx(file);
      try {
        await getDefaultStorage().updateContent(docId, trimExt(file.name), snapshot);
      } catch (err) {
        console.error('导入保存文档失败', err);
        useUiStore.getState().showToast(excelT('excel.toast.importSaveFailed'));
        return false;
      }
      // 图表/批注/超链接入队：编辑器重载挂载 Univer 后落位（队列在各解析模块，避免此处引入 Univer 依赖）
      queueImportedCharts(charts);
      queueImportedAnnotations(annotations);
      clearDraft(docId); // 防重载读到旧草稿
      void useDocumentsStore.getState().refresh().catch(() => {}); // 防列表刷新失败的 unhandledrejection
      useEditorStore.getState().bumpReload();
      // 超过 500KB 的文档宿主只存本地浏览器，导入后明确告知用户
      const localOnly = new TextEncoder().encode(JSON.stringify(snapshot ?? null)).length > LOCAL_ONLY_LIMIT_BYTES;
      const toast = localOnly
        ? charts.length
          ? excelT('excel.toast.importedWithChartsLocalOnly', { count: charts.length })
          : excelT('excel.toast.importedLocalOnly')
        : charts.length
          ? excelT('excel.toast.importedWithCharts', { count: charts.length })
          : excelT('excel.toast.imported');
      useUiStore.getState().showToast(toast);
      return true;
    } catch (err) {
      console.error('导入 Excel 失败', err);
      useUiStore.getState().showToast(excelT('excel.toast.importFailed'));
      return false;
    }
  } finally {
    useUiStore.getState().setImporting(false);
  }
}
