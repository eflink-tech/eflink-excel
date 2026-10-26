// 图表数据联动的纯逻辑：按源引用重建 explicit 数据（读值注入，供 node 单测）。
// Univer 交互（监听/重渲染/指纹调度）在 chartSettings.ts，本模块保持可独立测试。
import type { ExplicitChartData } from './chartOptions';
import type { LiveChartRefs } from '../xlsx/chartImport';

/** 二维选区值 → 扁平列（按列优先取首个非空列；单列引用常规形态） */
export function flattenRange(values: unknown[][] | null): unknown[] {
  if (!values?.length) return [];
  return values.map((row) => row[0] ?? '');
}

export const cellToNum = (v: unknown): number => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const n = Number.parseFloat(String(v ?? ''));
  return Number.isFinite(n) ? n : 0;
};

export const cellToStr = (v: unknown): string => {
  if (v == null) return '';
  if (typeof v === 'object') {
    const o = v as { result?: unknown; hyperlink?: string; text?: string; p?: { body?: { dataStream?: string } } };
    if (o.result != null) return String(o.result);
    if (typeof o.text === 'string') return o.text;
    if (typeof o.hyperlink === 'string') return o.hyperlink;
    // Univer 富文本单元格：dataStream 以 \r\n 结尾（段落分隔），还原为换行文本
    if (typeof o.p?.body?.dataStream === 'string') return o.p.body.dataStream.replace(/\r\n$/, '');
    return '';
  }
  return String(v);
};

/**
 * 按 liveRefs 从表格当前值重建 explicit。
 * 引用缺失的系列（numLit 字面量、筛选残留等）沿用 fallback 里同位置的旧数据；
 * 必需数据（类别/数值系列）完全缺失时返回 null——调用方保留旧图不动。
 * comboBars 标记组合图柱系列在前几位、折线在后。
 */
export function rebuildExplicitFromRefs(
  typeId: string,
  refs: LiveChartRefs & { comboBars?: number },
  read: (ref: string) => unknown[],
  fallback?: ExplicitChartData | null,
): ExplicitChartData | null {
  const fb = fallback ?? null;
  const fallbackCatSeries = fb?.series ?? [];
  const fallbackScatter = fb?.scatterSeries ?? [];
  const cats = refs.catRef ? read(refs.catRef).map(cellToStr) : fb?.categories;
  const nameSeries: { name: string; data: number[] }[] = [];
  const pointSeries: { name: string; data: [number, number][] }[] = [];
  let valIdx = 0;
  let ptIdx = 0;
  refs.series.forEach((r, i) => {
    const name = (r.nameRef ? cellToStr(read(r.nameRef)[0]) : '') || `系列${i + 1}`;
    if (r.valRef) {
      nameSeries.push({ name, data: read(r.valRef).map(cellToNum) });
      valIdx += 1;
    } else if (r.xRef && r.yRef) {
      const xs = read(r.xRef).map(cellToNum);
      const ys = read(r.yRef).map(cellToNum);
      pointSeries.push({ name, data: xs.map((x, j) => [x, ys[j] ?? 0]) });
      ptIdx += 1;
    } else if (fallbackCatSeries[valIdx]?.data) {
      // 无引用系列（字面量数据等）：沿用旧值，保持系列完整性
      nameSeries.push(fallbackCatSeries[valIdx]!);
      valIdx += 1;
    } else if (fallbackScatter[ptIdx]?.data) {
      pointSeries.push(fallbackScatter[ptIdx]!);
      ptIdx += 1;
    }
  });

  // 读空（引用失效/区域被删）视为缺数据：保留旧图而非渲染空图
  if (cats && !cats.length) return null;
  if (nameSeries.length && nameSeries.every((s) => !s.data.length)) return null;
  if (typeId === 'combo-bar-line') {
    const barCount = refs.comboBars ?? 1;
    const bars = nameSeries.slice(0, barCount);
    const lines = nameSeries.slice(barCount);
    if (!cats || !bars.length) return null;
    if (!lines.length && fb?.combo) lines.push(...fb.combo.lines);
    if (!lines.length) return null;
    return { categories: cats, combo: { bars, lines } };
  }
  if (typeId === 'scatter') {
    return pointSeries.length ? { scatterSeries: pointSeries } : null;
  }
  if (typeId === 'pie' || typeId === 'pie-doughnut') {
    if (!cats || !nameSeries.length) return null;
    return { pieData: cats.map((name, i) => ({ name, value: nameSeries[0]!.data[i] ?? 0 })) };
  }
  // 柱/折/条/面/雷达：类别 + 全部数值系列
  if (!cats || !nameSeries.length) return null;
  return { categories: cats, series: nameSeries };
}
