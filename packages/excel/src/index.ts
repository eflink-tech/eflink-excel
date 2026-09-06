/** @eflink-tech/excel 对外导出面：组件 + headless API + 存储适配器 */
import './styles.css';

// 组件
export { SheetEditor } from './components/SheetEditor';
export type { SheetEditorProps, SheetEditorBrand } from './components/SheetEditor';
export { ToastHost } from './components/ToastHost';
export { ConfirmDialogHost } from './components/ConfirmDialog';
export type { ConfirmOptions } from './store/uiStore';

// 存储适配器
export type { StorageAdapter } from './storage/types';
export { indexedDbStorage } from './storage/indexedDb';
export { memoryStorage } from './storage/memory';
export { getDefaultStorage, setDefaultStorage } from './storage/registry';
// 宿主外观：顶栏返回按钮链接
export { setEditorBackHref } from './core/chrome';

// headless：.efexcel 文档备份导入导出、xlsx 转换、PNG 截图导出
export { exportEfexcel, importEfexcel, DEFAULT_EFEXCEL_SOURCE } from './core/efexcel';
export type { EfexcelSource } from './core/efexcel';
export { importXlsx, exportXlsx, exportPng } from './core/xlsx/fileIO';

// 数据模型
export { createDocument, createEmptySnapshot, newId } from './types/spreadsheet';
export type {
  CellStyle,
  SnapshotCell,
  MergeRange,
  SnapshotSheet,
  WorkbookSnapshot,
  DocumentMeta,
  SheetDocument,
} from './types/spreadsheet';

// 状态（宿主可读取编辑器/文档/提示状态）
export { useEditorStore } from './store/editorStore';
export { useDocumentsStore } from './store/documentsStore';
export { useUiStore } from './store/uiStore';
