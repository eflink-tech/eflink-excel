// 图表设置元数据：插入图表时登记（类型/数据/元素/配色），双击图片后由右侧设置面板读取与修改，
// 修改通过 buildChartOption 重建配置并 setSource 原位替换图片（保持位置与尺寸）。
// 注意：元数据保存在内存中，页面刷新后图表仍显示但设置面板不可用（重新插入即可恢复）。
import { getUniverApi } from './controller';
import type { LiveChartRefs } from '../xlsx/chartImport';
import { flattenRange, rebuildExplicitFromRefs } from './chartRebuild';
export { rebuildExplicitFromRefs } from './chartRebuild';
import {
  buildChartOption,
  renderChartFrameToDataUrl,
  DEFAULT_CHART_ELEMENTS,
  type ChartElements,
  type ExplicitChartData,
} from './chartOptions';

export interface ChartMeta {
  drawingId: string;
  typeId: string;
  matrix: unknown[][];
  elements: ChartElements;
  paletteIndex: number;
  title: string;
  /** 插入时的数据范围 A1 记法（如 H1:I4），供面板「数据范围」展示与修改 */
  rangeRef: string;
  /** xlsx 导入图表的显式数据：重渲染时绕过 matrix 启发式（组合图/多系列散点等） */
  explicit?: ExplicitChartData;
  /** 图表所在 sheet 名（联动重读/图片查找按此定位，缺省视为当前活动表） */
  sheetName?: string;
  /** 各系列源数据引用：单元格编辑后按引用重读并自动重渲染（仅导入图表有） */
  liveRefs?: LiveChartRefs;
  /** 上次渲染的数据指纹：重读后值未变则跳过重渲染 */
  dataSignature?: string;
}

const chartMetas = new Map<string, ChartMeta>();

// 按工作簿 unitId 持久化到 localStorage：刷新页面后双击图表仍可打开设置面板
const STORAGE_PREFIX = 'eflink-excel:chart-meta:';
let currentUnitId: string | null = null;

/** 按工作簿（unitId）加载持久化的图表元数据，mountUniver 时调用 */
export function loadChartMetasForUnit(unitId: string): void {
  currentUnitId = unitId;
  chartMetas.clear();
  try {
    const raw = localStorage.getItem(STORAGE_PREFIX + unitId);
    if (!raw) return;
    for (const [id, meta] of Object.entries(JSON.parse(raw) as Record<string, ChartMeta>)) {
      // 合并默认值：兼容后续新增的元素开关字段
      chartMetas.set(id, {
        ...meta,
        elements: { ...DEFAULT_CHART_ELEMENTS, ...meta.elements },
      });
    }
    notifyMetaChange();
  } catch (err) {
    console.warn('[chartSettings] 读取图表元数据失败', err);
  }
}

function persistChartMetas(): void {
  if (!currentUnitId) return;
  try {
    localStorage.setItem(STORAGE_PREFIX + currentUnitId, JSON.stringify(Object.fromEntries(chartMetas)));
  } catch (err) {
    console.warn('[chartSettings] 保存图表元数据失败', err);
  }
}

// 元数据版本号 + 订阅：面板据此在元数据变化时重渲染（配合 useSyncExternalStore）
let metaVersion = 0;
const metaListeners = new Set<() => void>();

function notifyMetaChange(): void {
  metaVersion += 1;
  metaListeners.forEach((listener) => listener());
}

export function subscribeChartMeta(listener: () => void): () => void {
  metaListeners.add(listener);
  return () => metaListeners.delete(listener);
}

export function getChartMetaVersion(): number {
  return metaVersion;
}

export function setChartMeta(meta: ChartMeta): void {
  chartMetas.set(meta.drawingId, meta);
  notifyMetaChange();
  persistChartMetas();
}

export function getChartMeta(drawingId: string): ChartMeta | null {
  return chartMetas.get(drawingId) ?? null;
}

export function removeChartMeta(drawingId: string): void {
  chartMetas.delete(drawingId);
  notifyMetaChange();
  persistChartMetas();
}

/** 调试：列出全部元数据（DEV 调试句柄用） */
export function debugAllMetas(): [string, ChartMeta][] {
  return [...chartMetas.entries()];
}

export function defaultChartElements(): ChartElements {
  return { ...DEFAULT_CHART_ELEMENTS };
}

/** 按 A1 记法读取选区数据（如 H1:I4），无效范围返回 null */
export function getRangeMatrix(rangeRef: string): unknown[][] | null {
  const ws = getUniverApi()?.getActiveWorkbook()?.getActiveSheet();
  if (!ws) return null;
  try {
    const values = ws.getRange(rangeRef).getValues() as unknown[][];
    return values.length > 0 ? values : null;
  } catch {
    return null;
  }
}

// facade 类型从 API 推导（与 univer/importedCharts.ts 一致）
type UniverApiT = NonNullable<ReturnType<typeof getUniverApi>>;
type WorksheetT = NonNullable<ReturnType<NonNullable<ReturnType<UniverApiT['getActiveWorkbook']>>['getSheetByName']>>;

/** 按 meta 定位图表所在 sheet（联动图表可能在非活动 sheet） */
function sheetOfMeta(meta: ChartMeta): WorksheetT | null {
  const wb = getUniverApi()?.getActiveWorkbook();
  return (meta.sheetName ? wb?.getSheetByName(meta.sheetName) : wb?.getActiveSheet()) ?? null;
}

/** 按元数据重建图表并原位替换图片源（保持位置与尺寸） */
export async function refreshChart(drawingId: string): Promise<boolean> {
  const meta = chartMetas.get(drawingId);
  if (!meta) return false;
  const api = getUniverApi();
  const ws = sheetOfMeta(meta);
  const image = ws?.getImages().find((img) => img.getId() === drawingId);
  if (!api || !ws || !image) return false;
  const option = buildChartOption(meta.matrix, meta.typeId, false, meta, meta.explicit);
  const url = await renderChartFrameToDataUrl(option);
  return image.setSource(url);
}

/** 修改元数据并刷新图表 */
export async function updateChartMeta(drawingId: string, patch: Partial<ChartMeta>): Promise<boolean> {
  const meta = chartMetas.get(drawingId);
  if (!meta) return false;
  const next: ChartMeta = { ...meta, ...patch };
  // 显式更新 matrix（应用数据范围/行列互换）= 重新绑定表格数据，导入时的 explicit/源引用缓存作废
  if (patch.matrix) {
    delete next.explicit;
    delete next.liveRefs;
  }
  chartMetas.set(drawingId, next);
  notifyMetaChange();
  persistChartMetas();
  return refreshChart(drawingId);
}

// ---------- 数据联动：单元格编辑后按源引用重读并自动重渲染 ----------

/** 引用读取函数：A1 记法 → 该选区的二维值（注入以便纯逻辑单测） */
export type RefValueReader = (ws: unknown, ref: string) => unknown[][] | null;

/** 图表数据指纹：值变化才重渲染（ECharts 渲染 ~250ms，不能每次编辑都跑） */
function signatureOf(data: unknown): string {
  return JSON.stringify(data);
}

let liveRefreshTimer: ReturnType<typeof setTimeout> | null = null;

/**
 * 请求联动刷新（controller 在单元格/结构变更命令后调用）。
 * 防抖 700ms 合并连续编辑；逐图「重读 → 指纹比对 → 变了才重渲染」。
 */
export function requestLiveChartsRefresh(): void {
  if (liveRefreshTimer) clearTimeout(liveRefreshTimer);
  liveRefreshTimer = setTimeout(() => {
    liveRefreshTimer = null;
    void refreshLiveCharts();
  }, 700);
}

async function refreshLiveCharts(): Promise<void> {
  const api = getUniverApi();
  const wb = api?.getActiveWorkbook();
  if (!api || !wb) return;
  // 迭代前快照：refreshChart 为异步，期间用户删图会变更 Map，直接迭代器遍历可能中途抛错
  for (const [drawingId, meta] of [...chartMetas.entries()]) {
    try {
      const ws = sheetOfMeta(meta);
      if (!ws) continue;
      let rebuilt = false;
      if (meta.liveRefs) {
        const read = (ref: string) => flattenRange((ws.getRange(ref).getValues() ?? []) as unknown[][]);
        const explicit = rebuildExplicitFromRefs(meta.typeId, meta.liveRefs, read, meta.explicit);
        if (!explicit) continue;
        const sig = signatureOf(explicit);
        if (sig === meta.dataSignature) continue;
        meta.explicit = explicit;
        meta.dataSignature = sig;
        rebuilt = true;
      } else if (meta.rangeRef) {
        // 应用内插入的图表：按 rangeRef 重读矩阵
        const values = ws.getRange(meta.rangeRef).getValues() as unknown[][];
        if (!values?.length) continue;
        const sig = signatureOf(values);
        if (sig === meta.dataSignature) continue;
        meta.matrix = values;
        meta.dataSignature = sig;
        rebuilt = true;
      }
      if (rebuilt) {
        persistChartMetas();
        await refreshChart(drawingId);
      }
    } catch (err) {
      console.warn('[chartSettings] 联动刷新失败，跳过', drawingId, err);
    }
  }
}
