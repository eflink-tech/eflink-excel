// 图表设置元数据：插入图表时登记（类型/数据/元素/配色），双击图片后由右侧设置面板读取与修改，
// 修改通过 buildChartOption 重建配置并 setSource 原位替换图片（保持位置与尺寸）。
// 注意：元数据保存在内存中，页面刷新后图表仍显示但设置面板不可用（重新插入即可恢复）。
import { getUniverApi } from './controller';
import {
  buildChartOption,
  renderChartFrameToDataUrl,
  DEFAULT_CHART_ELEMENTS,
  type ChartElements,
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

/** 按元数据重建图表并原位替换图片源（保持位置与尺寸） */
export async function refreshChart(drawingId: string): Promise<boolean> {
  const meta = chartMetas.get(drawingId);
  if (!meta) return false;
  const api = getUniverApi();
  const ws = api?.getActiveWorkbook()?.getActiveSheet();
  const image = ws?.getImages().find((img) => img.getId() === drawingId);
  if (!api || !ws || !image) return false;
  const option = buildChartOption(meta.matrix, meta.typeId, false, meta);
  const url = await renderChartFrameToDataUrl(option);
  return image.setSource(url);
}

/** 修改元数据并刷新图表 */
export async function updateChartMeta(drawingId: string, patch: Partial<ChartMeta>): Promise<boolean> {
  const meta = chartMetas.get(drawingId);
  if (!meta) return false;
  chartMetas.set(drawingId, { ...meta, ...patch });
  notifyMetaChange();
  persistChartMetas();
  return refreshChart(drawingId);
}
