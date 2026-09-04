import {
  createUniver,
  defaultTheme,
  LocaleType,
  mergeLocales,
} from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import PresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN';
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import PresetSheetsFilterZhCN from '@univerjs/preset-sheets-filter/locales/zh-CN';
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort';
import PresetSheetsSortZhCN from '@univerjs/preset-sheets-sort/locales/zh-CN';
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing';
import PresetSheetsDrawingZhCN from '@univerjs/preset-sheets-drawing/locales/zh-CN';
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link';
import PresetSheetsHyperLinkZhCN from '@univerjs/preset-sheets-hyper-link/locales/zh-CN';
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace';
import PresetSheetsFindReplaceZhCN from '@univerjs/preset-sheets-find-replace/locales/zh-CN';
import { UniverSheetsThreadCommentPreset } from '@univerjs/preset-sheets-thread-comment';
import PresetSheetsThreadCommentZhCN from '@univerjs/preset-sheets-thread-comment/locales/zh-CN';
import '@univerjs/presets/lib/styles/preset-sheets-core.css';
import '@univerjs/presets/lib/styles/preset-sheets-filter.css';
import '@univerjs/presets/lib/styles/preset-sheets-sort.css';
import '@univerjs/presets/lib/styles/preset-sheets-drawing.css';
import '@univerjs/presets/lib/styles/preset-sheets-hyper-link.css';
import '@univerjs/presets/lib/styles/preset-sheets-find-replace.css';
import '@univerjs/presets/lib/styles/preset-sheets-thread-comment.css';
import { useUiStore } from '../../store/uiStore';
import {
  debugAllMetas,
  getChartMeta,
  loadChartMetasForUnit,
  removeChartMeta,
  setChartMeta,
  defaultChartElements,
} from './chartSettings';
import { buildChartOption, renderChartFrameToDataUrl } from './chartOptions';

/** 筛选/排序面板文案对齐企微表格（在 Univer zh-CN 基础上仅覆盖措辞，需展开合并避免丢键） */
const WeComZhCN = {
  'sheets-filter-ui': {
    ...PresetSheetsFilterZhCN['sheets-filter-ui'],
    panel: {
      ...PresetSheetsFilterZhCN['sheets-filter-ui'].panel,
      'by-values': '按选项',
      'select-all': '全部',
      confirm: '确定',
      'clear-filter': '重置',
      'search-placeholder': '搜索，支持用空格分隔多个关键词',
    },
  },
};
import type { WorkbookSnapshot } from '../../types/spreadsheet';

type UniverApi = ReturnType<typeof createUniver>['univerAPI'];

// Univer 实例的生命周期管理：一个页面同时只挂一个编辑器。
// API 通过模块级引用暴露给工具栏/公式栏（commands.ts 读取）。
let currentApi: UniverApi | null = null;

export function getUniverApi(): UniverApi | null {
  return currentApi;
}

export interface UniverController {
  getSnapshot(): WorkbookSnapshot | null;
  destroy(): void;
}

export interface MountHooks {
  /** 任意命令执行后触发（自动保存防抖入口） */
  onChange(): void;
}

export function mountUniver(
  container: HTMLElement,
  snapshot: WorkbookSnapshot,
  hooks: MountHooks,
): UniverController {
  const { univerAPI } = createUniver({
    locale: LocaleType.ZH_CN,
    // locales 必须按语言 ID 键控；mergeLocales 返回的是扁平语言包
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(
        PresetSheetsCoreZhCN,
        PresetSheetsFilterZhCN,
        PresetSheetsSortZhCN,
        PresetSheetsDrawingZhCN,
        PresetSheetsHyperLinkZhCN,
        PresetSheetsFindReplaceZhCN,
        PresetSheetsThreadCommentZhCN,
        WeComZhCN,
      ),
    },
    theme: defaultTheme,
    presets: [
      UniverSheetsCorePreset({
        container,
        // header=false 关闭 Univer 自带的 ribbon + 公式栏区域，
        // 工具栏与公式栏由 React 层（SheetToolbar / FormulaBar）按企微样式实现
        header: false,
      }),
      UniverSheetsFilterPreset(),
      UniverSheetsSortPreset(),
      UniverSheetsDrawingPreset(),
      UniverSheetsHyperLinkPreset(),
      UniverSheetsFindReplacePreset(),
      UniverSheetsThreadCommentPreset(),
    ],
  });
  currentApi = univerAPI;
  // 开发模式暴露调试句柄（生产构建不注入）
  if (import.meta.env.DEV) {
    (window as unknown as { __univer: unknown }).__univer = univerAPI;
    (window as unknown as { __chartDebug: unknown }).__chartDebug = {
      metas: debugAllMetas,
      lastSelected: () => lastSelectedChartId,
      openChartPanel: (drawingId: string) => useUiStore.getState().openChartPanel(drawingId),
      closeChartPanel: () => useUiStore.getState().closeChartPanel(),
      insertChartAt: async (rangeRef: string, typeId: string) => {
        const wb = univerAPI.getActiveWorkbook();
        const ws = wb?.getActiveSheet();
        const range = ws?.getRange(rangeRef);
        if (!wb || !ws || !range) return;
        const matrix = range.getValues() as unknown[][];
        const option = buildChartOption(matrix, typeId);
        const url = await renderChartFrameToDataUrl(option);
        await ws.insertImage(url, range.getColumn(), range.getRow());
        const images = ws.getImages();
        const added = images[images.length - 1];
        if (!added) return;
        await added.setSizeAsync(664, 424);
        setChartMeta({
          drawingId: added.getId(),
          typeId,
          matrix,
          elements: defaultChartElements(),
          paletteIndex: 0,
          title: '',
          rangeRef,
        });
        return added.getId();
      },
    };
  }

  univerAPI.createWorkbook(snapshot as never);

  // 恢复该工作簿持久化的图表元数据（双击图表打开设置面板依赖它）
  const workbookUnitId = univerAPI.getActiveWorkbook()?.getId();
  if (workbookUnitId) loadChartMetasForUnit(workbookUnitId);

  const stopWatch = univerAPI.onCommandExecuted(() => hooks.onChange());

  // 图表设置：记录选中的图表图片，双击时打开右侧设置面板；图片删除时同步清理元数据
  let lastSelectedChartId: string | null = null;
  interface SelectedFloatingImage {
    getId(): string;
    remove(): boolean;
  }
  let selectedFloatingImages: SelectedFloatingImage[] = [];
  const selectedWatch = univerAPI.addEvent(univerAPI.Event.OverGridImageSelected, (params) => {
    selectedFloatingImages = (params.selectedImages ?? []) as SelectedFloatingImage[];
    const image = selectedFloatingImages[0];
    lastSelectedChartId = image && getChartMeta(image.getId()) ? image.getId() : null;
  });
  const removedWatch = univerAPI.addEvent(univerAPI.Event.OverGridImageRemoved, (params) => {
    const ui = useUiStore.getState();
    for (const item of params.removeImageParams ?? []) {
      removeChartMeta(item.drawingId);
      if (ui.chartPanelDrawingId === item.drawingId) ui.closeChartPanel();
    }
  });
  /** 关闭 Univer 图片 Gallery 预览（补丁未命中时的兜底） */
  const dismissImagePreview = () => {
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  };
  const onCanvasDblClick = () => {
    if (!lastSelectedChartId) return;
    useUiStore.getState().openChartPanel(lastSelectedChartId);
    // Univer 内部 dblclick 可能同步弹出预览，下一微任务用 Escape 关掉
    queueMicrotask(dismissImagePreview);
  };
  container.addEventListener('dblclick', onCanvasDblClick, true);

  // 浮动图片/图表框：Del / Backspace 删除（Univer 内置快捷键依赖 FOCUSING_COMMON_DRAWINGS，自定义公式栏场景下需兜底）
  const onFloatingDeleteKey = (e: KeyboardEvent) => {
    if (e.key !== 'Delete' && e.key !== 'Backspace') return;
    const target = e.target as HTMLElement | null;
    if (target?.closest('input, textarea, select, [contenteditable="true"]')) return;

    const ws = univerAPI.getActiveWorkbook()?.getActiveSheet();
    if (!ws) return;

    const active = ws.getActiveImages() as SelectedFloatingImage[];
    const toDelete = active.length > 0 ? active : selectedFloatingImages;
    if (toDelete.length === 0) return;

    e.preventDefault();
    e.stopImmediatePropagation();
    ws.deleteImages(toDelete as never);
    selectedFloatingImages = [];
    lastSelectedChartId = null;
  };
  window.addEventListener('keydown', onFloatingDeleteKey, true);

  // 筛选面板打开时，把选区设为对应列的表头单元格：以选区高亮标识面板属于哪一列（企微样式）
  const filterPanelWatch = univerAPI.onCommandExecuted((command) => {
    if (command.id !== 'sheet.operation.open-filter-panel') return;
    const params = command.params as { unitId?: string; subUnitId?: string; col?: number } | undefined;
    const wb = univerAPI.getActiveWorkbook();
    if (!wb || !params || typeof params.col !== 'number') return;
    const ws = (params.subUnitId ? wb.getSheetBySheetId(params.subUnitId) : null) ?? wb.getActiveSheet();
    const filter = ws?.getFilter();
    if (!ws || !filter) return;
    ws.getRange(filter.getRange().getRow(), params.col).activate();
  });

  return {
    getSnapshot() {
      const workbook = univerAPI.getActiveWorkbook();
      if (!workbook) return null;
      return workbook.save() as never as WorkbookSnapshot;
    },
      destroy() {
        stopWatch.dispose();
        filterPanelWatch.dispose();
        selectedWatch.dispose();
        removedWatch.dispose();
        container.removeEventListener('dblclick', onCanvasDblClick, true);
        window.removeEventListener('keydown', onFloatingDeleteKey, true);
        univerAPI.dispose();
        if (currentApi === univerAPI) currentApi = null;
      },
  };
}
