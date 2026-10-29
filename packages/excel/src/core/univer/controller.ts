import {
  createUniver,
  defaultTheme,
  LocaleType,
  mergeLocales,
} from '@univerjs/presets';
import type { OfficeLocale } from '@eflink-tech/office-kit';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import PresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN';
import PresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import PresetSheetsCoreEsES from '@univerjs/preset-sheets-core/locales/es-ES';
import PresetSheetsCorePtBR from '@univerjs/preset-sheets-core/locales/pt-BR';
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import PresetSheetsFilterZhCN from '@univerjs/preset-sheets-filter/locales/zh-CN';
import PresetSheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US';
import PresetSheetsFilterEsES from '@univerjs/preset-sheets-filter/locales/es-ES';
import PresetSheetsFilterPtBR from '@univerjs/preset-sheets-filter/locales/pt-BR';
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort';
import PresetSheetsSortZhCN from '@univerjs/preset-sheets-sort/locales/zh-CN';
import PresetSheetsSortEnUS from '@univerjs/preset-sheets-sort/locales/en-US';
import PresetSheetsSortEsES from '@univerjs/preset-sheets-sort/locales/es-ES';
import PresetSheetsSortPtBR from '@univerjs/preset-sheets-sort/locales/pt-BR';
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing';
import PresetSheetsDrawingZhCN from '@univerjs/preset-sheets-drawing/locales/zh-CN';
import PresetSheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US';
import PresetSheetsDrawingEsES from '@univerjs/preset-sheets-drawing/locales/es-ES';
import PresetSheetsDrawingPtBR from '@univerjs/preset-sheets-drawing/locales/pt-BR';
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link';
import PresetSheetsHyperLinkZhCN from '@univerjs/preset-sheets-hyper-link/locales/zh-CN';
import PresetSheetsHyperLinkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US';
import PresetSheetsHyperLinkEsES from '@univerjs/preset-sheets-hyper-link/locales/es-ES';
import PresetSheetsHyperLinkPtBR from '@univerjs/preset-sheets-hyper-link/locales/pt-BR';
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace';
import PresetSheetsFindReplaceZhCN from '@univerjs/preset-sheets-find-replace/locales/zh-CN';
import PresetSheetsFindReplaceEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US';
import PresetSheetsFindReplaceEsES from '@univerjs/preset-sheets-find-replace/locales/es-ES';
import PresetSheetsFindReplacePtBR from '@univerjs/preset-sheets-find-replace/locales/pt-BR';
import { UniverSheetsThreadCommentPreset } from '@univerjs/preset-sheets-thread-comment';
import PresetSheetsThreadCommentZhCN from '@univerjs/preset-sheets-thread-comment/locales/zh-CN';
import PresetSheetsThreadCommentEnUS from '@univerjs/preset-sheets-thread-comment/locales/en-US';
import PresetSheetsThreadCommentEsES from '@univerjs/preset-sheets-thread-comment/locales/es-ES';
import PresetSheetsThreadCommentPtBR from '@univerjs/preset-sheets-thread-comment/locales/pt-BR';
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
  requestLiveChartsRefresh,
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

/** 界面语言 → Univer 内置语言包 */
const UNIVER_LOCALE_BY_OFFICE: Record<OfficeLocale, LocaleType> = {
  'zh-CN': LocaleType.ZH_CN,
  en: LocaleType.EN_US,
  es: LocaleType.ES_ES,
  'pt-BR': LocaleType.PT_BR,
};

function univerLocalesByOffice(locale: OfficeLocale) {
  switch (locale) {
    case 'en':
      return mergeLocales(
        PresetSheetsCoreEnUS,
        PresetSheetsFilterEnUS,
        PresetSheetsSortEnUS,
        PresetSheetsDrawingEnUS,
        PresetSheetsHyperLinkEnUS,
        PresetSheetsFindReplaceEnUS,
        PresetSheetsThreadCommentEnUS,
      );
    case 'es':
      return mergeLocales(
        PresetSheetsCoreEsES,
        PresetSheetsFilterEsES,
        PresetSheetsSortEsES,
        PresetSheetsDrawingEsES,
        PresetSheetsHyperLinkEsES,
        PresetSheetsFindReplaceEsES,
        PresetSheetsThreadCommentEsES,
      );
    case 'pt-BR':
      return mergeLocales(
        PresetSheetsCorePtBR,
        PresetSheetsFilterPtBR,
        PresetSheetsSortPtBR,
        PresetSheetsDrawingPtBR,
        PresetSheetsHyperLinkPtBR,
        PresetSheetsFindReplacePtBR,
        PresetSheetsThreadCommentPtBR,
      );
    default:
      return mergeLocales(
        PresetSheetsCoreZhCN,
        PresetSheetsFilterZhCN,
        PresetSheetsSortZhCN,
        PresetSheetsDrawingZhCN,
        PresetSheetsHyperLinkZhCN,
        PresetSheetsFindReplaceZhCN,
        PresetSheetsThreadCommentZhCN,
        WeComZhCN,
      );
  }
}

/** 预注册四种语言包，切换语言时无需重建 Univer 实例 */
const UNIVER_LOCALES: Record<string, ReturnType<typeof univerLocalesByOffice>> = {
  [LocaleType.ZH_CN]: univerLocalesByOffice('zh-CN'),
  [LocaleType.EN_US]: univerLocalesByOffice('en'),
  [LocaleType.ES_ES]: univerLocalesByOffice('es'),
  [LocaleType.PT_BR]: univerLocalesByOffice('pt-BR'),
};

type UniverApi = ReturnType<typeof createUniver>['univerAPI'];

// Univer 实例的生命周期管理：一个页面同时只挂一个编辑器。
// API 通过模块级引用暴露给工具栏/公式栏（commands.ts 读取）。
let currentApi: UniverApi | null = null;
let currentUniver: ReturnType<typeof createUniver>['univer'] | null = null;

export function getUniverApi(): UniverApi | null {
  return currentApi;
}

/** 暴露底层 DI 容器：导入批注需向 UserManagerService 注册原作者（facade 未开放 addUser） */
export function getUniverInjector(): unknown {
  const univer = currentUniver as unknown as { __getInjector?: () => unknown } | null;
  return univer?.__getInjector?.() ?? null;
}

/** 切换 Univer 内置界面语言（筛选面板、右键菜单等），不重建实例 */
export function setUniverLocale(locale: OfficeLocale): void {
  currentUniver?.setLocale(UNIVER_LOCALE_BY_OFFICE[locale]);
}

export interface UniverController {
  getSnapshot(): WorkbookSnapshot | null;
  destroy(): void;
}

export interface MountHooks {
  /** 任意命令执行后触发（置 dirty + 防抖写本地草稿入口） */
  onChange(): void;
}

export function mountUniver(
  container: HTMLElement,
  snapshot: WorkbookSnapshot,
  hooks: MountHooks,
  locale: OfficeLocale = 'zh-CN',
): UniverController {
  const { univer, univerAPI } = createUniver({
    locale: UNIVER_LOCALE_BY_OFFICE[locale],
    // locales 必须按语言 ID 键控；mergeLocales 返回的是扁平语言包
    locales: UNIVER_LOCALES,
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
  currentUniver = univer;
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

  // 数据联动：单元格赋值（编辑/粘贴/填充/清除）与行列增删后，按图表源引用重读并重渲染
  // （setSource 走 drawing 命令，不在集合内，不会循环触发）
  const DATA_MUTATION_RE = /^sheet\.command\.(set-range-values|insert-(row|col|multi-rows|multi-cols)|remove-(row|col|multi-rows|multi-cols))/;
  const liveChartWatch = univerAPI.onCommandExecuted((command) => {
    if (DATA_MUTATION_RE.test(command.id)) requestLiveChartsRefresh();
  });

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
        liveChartWatch.dispose();
        filterPanelWatch.dispose();
        selectedWatch.dispose();
        removedWatch.dispose();
        container.removeEventListener('dblclick', onCanvasDblClick, true);
        window.removeEventListener('keydown', onFloatingDeleteKey, true);
        univerAPI.dispose();
        if (currentApi === univerAPI) currentApi = null;
        if (currentUniver === univer) currentUniver = null;
      },
  };
}
