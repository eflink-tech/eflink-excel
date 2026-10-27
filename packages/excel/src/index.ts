/** @eflink-tech/excel 对外导出面：组件 + headless API + 存储适配器 */
// 顶栏/菜单栏样式来自 office-kit：宿主只引入 dist/styles.css，必须在此显式引入才会打进产物
import '@eflink-tech/office-kit/styles.css';
import './styles.css';
import { migrateLegacyLocalStorage } from './storage/legacy';

// 模块加载即把 efLink 命名空间的 localStorage 数据搬到 doculuma 命名空间（幂等、静默失败）
migrateLegacyLocalStorage();

// 组件
export { SheetEditor } from './components/SheetEditor';
export type { SheetEditorProps, SheetEditorBrand } from './components/SheetEditor';

// 品牌切换：'doculuma' | 'eflink' 内置品牌，或传对象自定义
export { builtinEditorBrand, resolveEditorBrandId, EDITOR_BRAND_IDS } from './brand';
export type { EditorBrandId } from './brand';
export { ToastHost } from './components/ToastHost';
export { ConfirmDialogHost } from './components/ConfirmDialog';
export type { ConfirmOptions } from './store/uiStore';

// 多语言：四套界面语言与编辑器专属文案
export { EXCEL_MESSAGES, setExcelLocale, getExcelLocale, excelT } from './i18n';
export type { OfficeLocale } from '@eflink-tech/office-kit';

// 存储适配器
export type { StorageAdapter } from './storage/types';
export { indexedDbStorage } from './storage/indexedDb';
export { memoryStorage } from './storage/memory';
export { getDefaultStorage, setDefaultStorage } from './storage/registry';
// 本地存储键名（唯一来源）与旧命名空间迁移
export {
  CHART_META_PREFIX,
  DRAFT_PREFIX,
  EXCEL_AI_DB_NAME,
  EXCEL_DB_NAME,
  LAST_DOC_KEY,
  RECENT_COLORS_KEYS,
} from './storage/keys';
export { migrateLegacyLocalStorage } from './storage/legacy';
// 宿主外观：顶栏返回按钮链接
export { setEditorBackHref } from './core/chrome';

// 分享桥接：宿主注入分享实现（编辑器分享按钮调用）
export { setExcelShareHandler } from './core/share/shareBridge';
export type { ExcelShareHandler, ExcelShareDoc, ExcelShareResult } from './core/share/shareBridge';

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
