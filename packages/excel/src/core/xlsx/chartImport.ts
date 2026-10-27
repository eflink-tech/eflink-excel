// xlsx 图表导入（解析侧）：exceljs 完全不解析图表，这里直接读 zip 里的
// workbook→sheet→drawing→chart 关系链，把每张图表还原为「锚点 + 类型 + 缓存数据」。
// 渲染成图片与插入编辑器分别在 chartOptions / univer/importedCharts 完成；
// 本模块保持纯逻辑（不依赖 DOM/Univer），供 node 环境单测。
import JSZip from 'jszip';
import { SaxesParser } from 'saxes';
import type { ExplicitChartData } from '../univer/chartOptions';

/** EMU（Office 角度单位）→ px */
const EMU_TO_PX = 1 / 9525;

/** 浮动锚点：from 为 0 基行列；twoCellAnchor 有 to（按列宽行高量尺寸），oneCellAnchor 有 ext（直接像素） */
export interface ChartAnchor {
  col: number;
  row: number;
  colOffPx: number;
  rowOffPx: number;
  toCol?: number;
  toRow?: number;
  toColOffPx?: number;
  toRowOffPx?: number;
  extWidthPx?: number;
  extHeightPx?: number;
}

/** 图表源数据引用（A1 记法、表内相对；编辑联动时按引用重读单元格值，见 chartSettings.requestLiveChartsRefresh） */
export interface LiveChartRefs {
  /** 类别轴引用 */
  catRef?: string;
  /** 每系列的数值/坐标引用（与 explicit 系列一一对应） */
  series: { nameRef?: string; valRef?: string; xRef?: string; yRef?: string }[];
  /** 组合图（combo-bar-line）柱系列数：series 前段为柱、后段为折线 */
  comboBars?: number;
}

/** 一张待插入图表的全部信息（dataUrl 由渲染阶段补充） */
export interface ImportedChart {
  sheetName: string;
  anchor: ChartAnchor;
  /** chartOptions 的类型 id；柱+折线组合图为 'combo-bar-line' */
  typeId: string;
  title?: string;
  explicit: ExplicitChartData;
  /** 首系列数据区引用（A1 记法，供设置面板「数据范围」展示） */
  rangeRef: string;
  /** 各系列源数据引用：编辑单元格后按此重读并自动重渲染 */
  liveRefs?: LiveChartRefs;
  dataUrl?: string;
}

// ---------- 轻量 XML DOM（node 环境无 DOMParser，用 saxes 建树；仅取 localName，命名空间前缀不参与匹配） ----------

interface XNode {
  name: string;
  attrs: Record<string, string>;
  children: XNode[];
  text: string;
}

function parseXml(xml: string): XNode | null {
  const parser = new SaxesParser({ fragment: false });
  let root: XNode | null = null;
  let failed = false;
  const stack: XNode[] = [];
  parser.on('error', () => {
    failed = true;
  });
  parser.on('opentag', (node) => {
    const el: XNode = { name: node.name.includes(':') ? node.name.split(':').pop()! : node.name, attrs: {}, children: [], text: '' };
    for (const [k, v] of Object.entries(node.attributes)) {
      el.attrs[k.includes(':') ? k.split(':').pop()! : k] = String(v);
    }
    const parent = stack[stack.length - 1];
    if (parent) parent.children.push(el);
    else if (!root) root = el;
    stack.push(el);
  });
  parser.on('text', (t) => {
    const top = stack[stack.length - 1];
    if (top) top.text += t;
  });
  parser.on('closetag', () => {
    stack.pop();
  });
  parser.write(xml).close();
  return failed ? null : root;
}

function findFirst(node: XNode | null, name: string): XNode | null {
  if (!node) return null;
  if (node.name === name) return node;
  for (const child of node.children) {
    const hit = findFirst(child, name);
    if (hit) return hit;
  }
  return null;
}

function findAll(node: XNode | null, name: string, out: XNode[] = []): XNode[] {
  if (!node) return out;
  if (node.name === name) out.push(node);
  for (const child of node.children) findAll(child, name, out);
  return out;
}

/** 取首个后代指定标签的文本 */
function findText(node: XNode | null, name: string): string | undefined {
  return findFirst(node, name)?.text.trim() || undefined;
}

/** 取首个后代指定标签的 val 属性（OOXML 枚举值惯用属性而非元素文本，如 <c:grouping val="clustered"/>） */
function findVal(node: XNode | null, name: string): string | undefined {
  return findFirst(node, name)?.attrs.val || undefined;
}

// ---------- chart XML → 类型 + 数据 ----------

interface RawSer {
  name: string;
  cats: string[];
  vals: number[];
  xs: number[];
  ys: number[];
  catRef?: string;
  nameRef?: string;
  valRef?: string;
  xRef?: string;
  yRef?: string;
}

/** 类别/数值缓存：c:pt 按 idx 对齐成稠密数组（multiLvlStrRef 取最深层 lvl） */
function cacheStrings(refEl: XNode | null): string[] {
  if (!refEl) return [];
  const lvls = findAll(refEl, 'lvl');
  const container = lvls.length ? lvls[lvls.length - 1] : refEl;
  const pts = findAll(container, 'pt');
  if (!pts.length) return [];
  const slots: string[] = [];
  for (const pt of pts) {
    const idx = Number(pt.attrs.idx ?? slots.length);
    slots[idx] = (findFirst(pt, 'v')?.text ?? '').trim();
  }
  return Array.from(slots, (s) => s ?? '');
}

function cacheNumbers(refEl: XNode | null): number[] {
  return cacheStrings(refEl).map((s) => {
    const n = Number.parseFloat(s);
    return Number.isFinite(n) ? n : 0;
  });
}

/** c:f 引用（如 `常用工具!$C$3:$C$9`）→ 表内 A1 记法（C3:C9） */
function refOf(refEl: XNode | null): string | undefined {
  const f = findText(refEl, 'f');
  return f ? f.replace(/^[^!]*!/, '').replace(/\$/g, '') || undefined : undefined;
}

/** 收集图表节点有效系列：c:filteredBarSeries/filteredLineSeries 等容器里是图表筛选器
 * 隐藏的系列（Excel 取消勾选后数据仍保留在 XML），渲染与联动都不应参与 */
function collectSers(node: XNode): XNode[] {
  const out: XNode[] = [];
  const walk = (el: XNode) => {
    for (const child of el.children) {
      if (child.name.startsWith('filtered')) continue;
      if (child.name === 'ser') out.push(child);
      else walk(child);
    }
  };
  walk(node);
  return out;
}

function parseSer(ser: XNode): RawSer {
  const idxAttr = findAll(ser, 'idx')[0]?.attrs.idx;
  const fallbackNo = idxAttr != null ? Number(idxAttr) + 1 : 1;
  const tx = findFirst(ser, 'tx');
  const cat = findFirst(ser, 'cat');
  const val = findFirst(ser, 'val');
  const xVal = findFirst(ser, 'xVal');
  const yVal = findFirst(ser, 'yVal');
  return {
    name: findText(tx, 'v') ?? `系列${fallbackNo}`,
    cats: cacheStrings(cat),
    vals: cacheNumbers(val),
    xs: cacheNumbers(xVal),
    ys: cacheNumbers(yVal),
    nameRef: refOf(tx),
    catRef: refOf(cat),
    valRef: refOf(val),
    xRef: refOf(xVal),
    yRef: refOf(yVal),
  };
}

/** 组图分组 → typeId（percentStacked 的横向条形无对应变体，按堆积处理） */
function barTypeId(dir: string, grouping: string): string {
  if (dir === 'bar') return grouping === 'clustered' ? 'bar-horizontal' : 'bar-horizontal-stack';
  if (grouping === 'stacked' || grouping === 'percentStacked') return grouping === 'percentStacked' ? 'bar-percent' : 'bar-stack';
  return 'bar-cluster';
}

function lineTypeId(grouping: string, hasMarker: boolean): string {
  if (grouping === 'stacked') return 'line-stack';
  return hasMarker ? 'line-marker' : 'line';
}

/** chart1.xml 根节点内容 → 导入描述；结构不认识时返回 null（跳过该图） */
export function parseChartXml(xml: string): Pick<ImportedChart, 'typeId' | 'title' | 'explicit' | 'rangeRef' | 'liveRefs'> | null {
  const root = parseXml(xml);
  const chart = findFirst(root, 'chart');
  if (!chart) return null;
  const plotArea = findFirst(chart, 'plotArea');
  if (!plotArea) return null;

  const autoDeleted = findFirst(chart, 'autoTitleDeleted')?.attrs.val === '1';
  const title = autoDeleted ? undefined : findAll(findFirst(chart, 'title'), 't').map((t) => t.text).join('').trim() || undefined;

  // plotArea 里以 Chart 结尾的子节点即图表本体（catAx/valAx 等轴节点被排除）
  const chartNodes = plotArea.children.filter((c) => c.name.endsWith('Chart'));
  if (!chartNodes.length) return null;
  const nodeTypes = chartNodes.map((n) => n.name);
  const sersByNode = chartNodes.map((n) => collectSers(n).map(parseSer));
  const allSers = sersByNode.flat();
  if (!allSers.length) return null;

  const firstSer = allSers[0]!;
  const firstSerEl = findFirst(chartNodes[0]!, 'ser');
  const dataRefEl =
    findFirst(firstSerEl, 'cat') ?? findFirst(firstSerEl, 'val') ?? findFirst(firstSerEl, 'xVal') ?? findFirst(firstSerEl, 'yVal') ?? firstSerEl;
  const rangeRef = refOf(dataRefEl) ?? '';
  // 每系列源引用（编辑联动的重读依据）；任一系列缺引用则整体不启用联动（宁可静态也不错位）
  const seriesRefs = allSers.map((s) => ({ nameRef: s.nameRef, valRef: s.valRef, xRef: s.xRef, yRef: s.yRef }));
  const barIdx = nodeTypes.findIndex((t) => /^bar(\dD)?Chart$/.test(t));
  const lineIdx = nodeTypes.indexOf('lineChart');
  const hasBar = barIdx >= 0;
  const hasLine = lineIdx >= 0;
  const hasCombo = hasBar && hasLine && sersByNode[barIdx]!.length > 0 && sersByNode[lineIdx]!.length > 0;
  const liveRefs =
    (firstSer.catRef || firstSer.valRef || firstSer.xRef) && seriesRefs.some((r) => r.valRef || r.xRef)
      ? {
          catRef: firstSer.catRef,
          series: seriesRefs,
          ...(hasCombo ? { comboBars: sersByNode[barIdx]!.length } : {}),
        }
      : undefined;

  // 柱 + 折线组合（帕累托图等）：柱系列走主轴、折线系列走次轴
  if (hasBar && hasLine) {
    const bars = sersByNode[barIdx] ?? [];
    const lines = sersByNode[lineIdx] ?? [];
    if (bars.length && lines.length) {
      return {
        typeId: 'combo-bar-line',
        title,
        rangeRef,
        liveRefs,
        explicit: {
          categories: (bars[0] ?? lines[0]!).cats,
          combo: { bars: bars.map(toSeries), lines: lines.map(toSeries) },
        },
      };
    }
  }

  const node = chartNodes[0]!;
  const sers = sersByNode[0]!;
  const grouping = findVal(node, 'grouping') ?? 'clustered';
  const t = node.name;
  const series = sers.map(toSeries);
  const categories = firstSer.cats;

  if (/^bar(\dD)?Chart$/.test(t)) {
    const dir = findVal(node, 'barDir') ?? 'col';
    return { typeId: barTypeId(dir, grouping), title, rangeRef, liveRefs, explicit: { categories, series } };
  }
  if (t === 'lineChart' || t === 'areaChart' || t === 'line3DChart') {
    // 数据标记按系列粒度声明（c:ser/c:marker/c:symbol），任一系列带标记即按带标记折线处理
    const serEls = collectSers(node);
    const hasMarker = serEls.some((serEl) => (findVal(findFirst(serEl, 'marker'), 'symbol') ?? 'none') !== 'none');
    const typeId = t === 'areaChart' ? (grouping === 'stacked' ? 'area-stack' : 'area') : lineTypeId(grouping, hasMarker);
    return { typeId, title, rangeRef, liveRefs, explicit: { categories, series } };
  }
  if (t === 'radarChart') return { typeId: 'radar', title, rangeRef, liveRefs, explicit: { categories, series } };
  if (t === 'scatterChart') {
    return {
      typeId: 'scatter',
      title,
      rangeRef,
      liveRefs,
      explicit: { scatterSeries: sers.map((s) => ({ name: s.name, data: s.xs.map((x, i) => [x, s.ys[i] ?? 0] as [number, number]) })) },
    };
  }
  if (t === 'pieChart' || t === 'pie3DChart' || t === 'doughnutChart') {
    return {
      typeId: t === 'doughnutChart' ? 'pie-doughnut' : 'pie',
      title,
      rangeRef,
      liveRefs,
      explicit: { pieData: firstSer.cats.map((name, i) => ({ name, value: firstSer.vals[i] ?? 0 })) },
    };
  }
  return null;
}

function toSeries(s: RawSer): { name: string; data: number[] } {
  return { name: s.name, data: s.vals };
}

// ---------- 关系链遍历：workbook → sheet → drawing → chart ----------

function dirname(path: string): string {
  const i = path.lastIndexOf('/');
  return i === -1 ? '' : path.slice(0, i);
}

/** rels 里的相对路径 → zip 内绝对路径（以 / 开头的包绝对路径也归一化） */
function resolvePath(fromDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (fromDir ? fromDir.split('/') : []).concat(target.split('/'));
  const out: string[] = [];
  for (const p of parts) {
    if (!p || p === '.') continue;
    if (p === '..') out.pop();
    else out.push(p);
  }
  return out.join('/');
}

/** 解析 *_rels/*.rels 内容为 Id → Target 映射（Relationship 属性名首字母大写） */
async function readRels(zip: JSZip, relsPath: string): Promise<Record<string, string>> {
  const file = zip.file(relsPath);
  if (!file) return {};
  const root = parseXml(await file.async('string'));
  const map: Record<string, string> = {};
  for (const rel of findAll(root, 'Relationship')) {
    map[rel.attrs.Id ?? rel.attrs.id ?? ''] = rel.attrs.Target ?? rel.attrs.target ?? '';
  }
  return map;
}

/** drawingN.xml：图表锚点 + 对应 chart xml 路径（非图表锚点忽略） */
async function readDrawingAnchors(zip: JSZip, drawingPath: string): Promise<{ anchor: ChartAnchor; chartPath: string }[]> {
  const root = parseXml(await zip.file(drawingPath)!.async('string'));
  const rels = await readRels(zip, `${dirname(drawingPath)}/_rels/${drawingPath.split('/').pop()}.rels`);
  /** xdr:from / xdr:to 单元格锚 → 行列号（0 基）与 EMU 偏移换算的像素 */
  const readCell = (el: XNode | null) => ({
    col: Number(findText(el, 'col') ?? 0),
    row: Number(findText(el, 'row') ?? 0),
    colOffPx: Math.round(Number(findText(el, 'colOff') ?? 0) * EMU_TO_PX),
    rowOffPx: Math.round(Number(findText(el, 'rowOff') ?? 0) * EMU_TO_PX),
  });
  const results: { anchor: ChartAnchor; chartPath: string }[] = [];
  for (const anchorEl of root ? [...findAll(root, 'twoCellAnchor'), ...findAll(root, 'oneCellAnchor')] : []) {
    const chartRef = findFirst(anchorEl, 'chart');
    const relTarget = chartRef?.attrs.id ? rels[chartRef.attrs.id] : undefined;
    if (!relTarget) continue;

    const anchor = { ...readCell(findFirst(anchorEl, 'from')) } as ChartAnchor;
    const to = findFirst(anchorEl, 'to');
    if (to) {
      const t = readCell(to);
      anchor.toCol = t.col;
      anchor.toRow = t.row;
      anchor.toColOffPx = t.colOffPx;
      anchor.toRowOffPx = t.rowOffPx;
    } else {
      // oneCellAnchor：xdr:ext 直接给 EMU 尺寸
      const ext = findFirst(anchorEl, 'ext');
      anchor.extWidthPx = Math.round(Number(ext?.attrs.cx ?? 0) * EMU_TO_PX);
      anchor.extHeightPx = Math.round(Number(ext?.attrs.cy ?? 0) * EMU_TO_PX);
    }
    results.push({ anchor, chartPath: resolvePath(dirname(drawingPath), relTarget) });
  }
  return results;
}

/**
 * 从 xlsx 原始字节提取全部图表（含隐藏 sheet 上的，由调用方决定是否落位）。
 * 任何一环缺失（无 drawing/rels/图表 XML 不认识）都静默跳过，绝不影响单元格导入。
 */
export async function extractCharts(buffer: ArrayBuffer): Promise<ImportedChart[]> {
  const zip = await JSZip.loadAsync(buffer);
  const out: ImportedChart[] = [];

  const workbookXml = parseXml((await zip.file('xl/workbook.xml')?.async('string')) ?? '');
  const workbookRels = await readRels(zip, 'xl/_rels/workbook.xml.rels');
  for (const sheetEl of workbookXml ? findAll(workbookXml, 'sheet') : []) {
    const sheetName = sheetEl.attrs.name ?? '';
    // 转义的工作表名（含空格等）在 rels target 上不体现，直接用即可
    const sheetPath = resolvePath('xl', workbookRels[sheetEl.attrs.id ?? ''] ?? '');
    if (!sheetPath || !zip.file(sheetPath)) continue;
    const sheetRelsPath = `${dirname(sheetPath)}/_rels/${sheetPath.split('/').pop()}.rels`;
    const sheetRels = await readRels(zip, sheetRelsPath);
    const drawingTarget = Object.values(sheetRels).find((t) => /drawings\/drawing\d+\.xml$/.test(t));
    if (!drawingTarget) continue;
    const drawingPath = resolvePath(dirname(sheetPath), drawingTarget);
    if (!zip.file(drawingPath)) continue;

    for (const { anchor, chartPath } of await readDrawingAnchors(zip, drawingPath)) {
      const chartXml = await zip.file(chartPath)?.async('string');
      if (!chartXml) continue;
      const parsed = parseChartXml(chartXml);
      if (!parsed) continue;
      out.push({ sheetName, anchor, ...parsed });
    }
  }
  return out;
}

// ---------- 待插入队列（渲染后的图片在编辑器挂载完成后落位，见 univer/importedCharts.ts） ----------

let pendingCharts: ImportedChart[] = [];

/** 渲染完成待插入的图表入队（导入动作调用；清空旧队列防串档） */
export function queueImportedCharts(charts: ImportedChart[]): void {
  pendingCharts = charts;
}

/** 取走队列（一次性消费；编辑器每次挂载后调用） */
export function takeImportedCharts(): ImportedChart[] {
  const charts = pendingCharts;
  pendingCharts = [];
  return charts;
}

/** 窥视队首（不移除）；配合 shiftImportedChart 逐张消费，实例换代时剩余图表整体留在队列 */
export function peekImportedChart(): ImportedChart | null {
  return pendingCharts[0] ?? null;
}

/** 取走队首一张 */
export function shiftImportedChart(): ImportedChart | null {
  return pendingCharts.shift() ?? null;
}

/** 塞回队首：落位中途实例被销毁时退回，等下一次挂载的消费者接手 */
export function unshiftImportedChart(chart: ImportedChart): void {
  pendingCharts.unshift(chart);
}
