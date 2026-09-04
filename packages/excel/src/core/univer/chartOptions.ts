/**
 * 图表类型注册表与 ECharts 配置构建。
 * 工具栏的图表选择面板（缩略图预览）与插入命令共用同一份类型定义和构建逻辑，
 * 保证「所见即所插」。选区数据约定：柱状/折线类首行为系列名、首列为类别名；
 * 饼图/散点图取第一、二列；雷达图首列为类别名（指标名）。
 */
import * as echarts from 'echarts';
import type { EChartsOption } from 'echarts';

export interface ChartTypeSpec {
  id: string;
  label: string;
}

export interface ChartGroup {
  group: string;
  types: ChartTypeSpec[];
}

/** 图表类型注册表（顺序即面板展示顺序） */
export const CHART_GROUPS: ChartGroup[] = [
  {
    group: '柱状图',
    types: [
      { id: 'bar-cluster', label: '簇状柱状图' },
      { id: 'bar-stack', label: '堆积柱状图' },
      { id: 'bar-percent', label: '百分比堆积柱状图' },
    ],
  },
  {
    group: '折线图',
    types: [
      { id: 'line', label: '折线图' },
      { id: 'line-stack', label: '堆积折线图' },
      { id: 'line-marker', label: '带数据标记的折线图' },
    ],
  },
  {
    group: '饼图',
    types: [
      { id: 'pie', label: '饼图' },
      { id: 'pie-doughnut', label: '圆环图' },
    ],
  },
  {
    group: '条形图',
    types: [
      { id: 'bar-horizontal', label: '条形图' },
      { id: 'bar-horizontal-stack', label: '堆积条形图' },
    ],
  },
  {
    group: '面积图',
    types: [
      { id: 'area', label: '面积图' },
      { id: 'area-stack', label: '堆积面积图' },
    ],
  },
  {
    group: '散点图',
    types: [{ id: 'scatter', label: '散点图' }],
  },
  {
    group: '雷达图',
    types: [{ id: 'radar', label: '雷达图' }],
  },
];

/** 图表选择面板缩略图的示例数据 */
export const SAMPLE_CHART_MATRIX: unknown[][] = [
  ['月份', '销售额', '利润'],
  ['一月', 120, 90],
  ['二月', 210, 150],
  ['三月', 170, 160],
];

/** 图表元素开关（对应图表设置面板的「图表元素」） */
export interface ChartElements {
  legend: boolean; // 图例
  gridLine: boolean; // 网格线
  axis: boolean; // 坐标轴
  trendline: boolean; // 趋势线（柱状/折线类，最小二乘拟合）
  chartTitle: boolean; // 图表标题
  dataLabel: boolean; // 数据标签
  axisTitle: boolean; // 轴标题
}

export const DEFAULT_CHART_ELEMENTS: ChartElements = {
  legend: true,
  gridLine: true,
  axis: true,
  trendline: false,
  chartTitle: false,
  dataLabel: false,
  axisTitle: false,
};

/** 图表配色方案（与「图表颜色」下拉对应） */
export const CHART_PALETTES: string[][] = [
  ['#5b8ff9', '#61ddaa', '#f6bd16', '#7262fd', '#78d3f8', '#9661bc'],
  ['#3370ff', '#34c724', '#ff9f1a', '#f5556a', '#8a5cf6', '#00b5d8'],
  ['#1f4e79', '#2e75b6', '#9dc3e6', '#ffd966', '#c55a11', '#70ad47'],
  ['#334155', '#64748b', '#0ea5e9', '#22c55e', '#f59e0b', '#ef4444'],
];

/** 线性趋势线（最小二乘） */
function trendData(data: number[]): number[] {
  const n = data.length;
  if (n < 2) return [...data];
  let sx = 0;
  let sy = 0;
  let sxy = 0;
  let sxx = 0;
  data.forEach((y, x) => {
    sx += x;
    sy += y;
    sxy += x * y;
    sxx += x * x;
  });
  const denom = n * sxx - sx * sx;
  if (!denom) return [...data];
  const k = (n * sxy - sx * sy) / denom;
  const b = (sy - k * sx) / n;
  return data.map((_, x) => +(k * x + b).toFixed(2));
}

/** 渲染图表选项为带边框阴影的 PNG dataURL（白色底板 + 细边框 + 投影） */
export async function renderChartFrameToDataUrl(option: EChartsOption): Promise<string> {
  const div = document.createElement('div');
  div.style.cssText = 'position:fixed;left:-99999px;top:0;width:640px;height:400px;';
  document.body.appendChild(div);
  const chart = echarts.init(div);
  let chartUrl: string;
  try {
    chart.setOption(option);
    await new Promise((resolve) => setTimeout(resolve, 200));
    chartUrl = chart.getDataURL({ pixelRatio: 2, backgroundColor: 'transparent' });
  } finally {
    chart.dispose();
    div.remove();
  }

  const image = new Image();
  await new Promise<void>((resolve, reject) => {
    image.onload = () => resolve();
    image.onerror = () => reject(new Error('图表图片加载失败'));
    image.src = chartUrl;
  });
  const pad = 24; // 2 倍像素比下的 12 逻辑像素边距
  const canvas = document.createElement('canvas');
  canvas.width = image.width + pad * 2;
  canvas.height = image.height + pad * 2;
  const ctx = canvas.getContext('2d');
  if (!ctx) return chartUrl;
  ctx.shadowColor = 'rgba(15, 23, 42, 0.18)';
  ctx.shadowBlur = 20;
  ctx.shadowOffsetY = 6;
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(pad, pad, image.width, image.height);
  ctx.shadowColor = 'transparent';
  ctx.strokeStyle = 'rgba(31, 41, 55, 0.16)';
  ctx.lineWidth = 2;
  ctx.strokeRect(pad + 1, pad + 1, image.width - 2, image.height - 2);
  ctx.drawImage(image, pad, pad);
  return canvas.toDataURL('image/png');
}

function toNumber(value: unknown): number {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const parsed = Number.parseFloat(String(value ?? ''));
  return Number.isFinite(parsed) ? parsed : 0;
}

function isNumericLike(value: unknown): boolean {
  if (typeof value === 'number') return true;
  return value !== null && value !== '' && Number.isFinite(Number.parseFloat(String(value)));
}

interface CategoryData {
  categories: string[];
  series: { name: string; data: number[] }[];
}

/** 类别类图表（柱/折/条/面积/雷达）的数据提取 */
function extractCategoryData(matrix: unknown[][]): CategoryData {
  const height = matrix.length;
  const width = matrix[0]?.length ?? 0;
  const hasHeaderRow = height > 1 && !matrix[0]!.some((v) => isNumericLike(v) && v !== '');
  const hasHeaderCol = width > 1 && matrix.slice(1).every((row) => !isNumericLike(row[0]));
  const rowStart = hasHeaderRow ? 1 : 0;
  const colStart = hasHeaderCol ? 1 : 0;

  const categories: string[] = [];
  for (let r = rowStart; r < height; r++) {
    categories.push(hasHeaderCol ? String(matrix[r]?.[0] ?? `行${r + 1}`) : `行${r + 1}`);
  }
  const series: { name: string; data: number[] }[] = [];
  for (let c = colStart; c < width; c++) {
    const name = hasHeaderRow ? String(matrix[0]?.[c] ?? `列${c + 1}`) : `列${c + 1}`;
    const data: number[] = [];
    for (let r = rowStart; r < height; r++) data.push(toNumber(matrix[r]?.[c]));
    series.push({ name, data });
  }
  if (series.length === 0 && height > 0) {
    series.push({ name: '系列1', data: matrix.map((row) => toNumber(row[0])) });
  }
  return { categories, series };
}

/** 饼图数据提取：第一列名称 + 第二列数值（单列时行号作名称） */
function extractPieData(matrix: unknown[][]): { name: string; value: number }[] {
  const height = matrix.length;
  const width = matrix[0]?.length ?? 0;
  const hasHeaderRow = height > 1 && !matrix[0]!.some((v) => isNumericLike(v) && v !== '');
  const dataStart = hasHeaderRow && width >= 2 ? 1 : 0;
  const pairs: { name: string; value: number }[] = [];
  for (let r = dataStart; r < height; r++) {
    const value = toNumber(width >= 2 ? matrix[r]?.[1] : matrix[r]?.[0]);
    const name = width >= 2 ? String(matrix[r]?.[0] ?? `项${r + 1}`) : `行${r + 1}`;
    if (name) pairs.push({ name, value });
  }
  return pairs;
}

/** 散点图数据提取：第一列 x + 第二列 y */
function extractScatterData(matrix: unknown[][]): { series: { name: string; data: [number, number][] }[] } {
  const height = matrix.length;
  const width = matrix[0]?.length ?? 0;
  const hasHeaderRow = height > 1 && !matrix[0]!.some((v) => isNumericLike(v) && v !== '');
  const rowStart = hasHeaderRow ? 1 : 0;
  const data: [number, number][] = [];
  for (let r = rowStart; r < height; r++) {
    const x = toNumber(matrix[r]?.[0]);
    const y = toNumber(matrix[r]?.[1]);
    data.push([x, y]);
  }
  const name = hasHeaderRow && width >= 2 ? String(matrix[0]?.[1]) : '散点图';
  return { series: [{ name, data }] };
}

/**
 * 构建图表 ECharts 配置。
 * @param matrix 选区（或示例）二维数据
 * @param typeId CHART_GROUPS 中注册的类型 id
 * @param mini 缩略图模式：压缩边距、隐藏图例、缩小字号
 * @param meta 图表设置面板的元信息（元素开关 / 配色 / 标题）
 */
export function buildChartOption(
  matrix: unknown[][],
  typeId: string,
  mini = false,
  meta?: { elements?: Partial<ChartElements>; paletteIndex?: number; title?: string },
): EChartsOption {
  const el: ChartElements = { ...DEFAULT_CHART_ELEMENTS, ...meta?.elements };
  const palette = CHART_PALETTES[Math.min(meta?.paletteIndex ?? 0, CHART_PALETTES.length - 1)];
  const base: EChartsOption = {
    animation: false,
    backgroundColor: 'transparent',
    color: palette,
  };

  // echarts 6 图例文字默认颜色在透明背景导出时不可见，需显式给色
  const textStyle = { color: '#1f2329', ...(mini ? { fontSize: 8 } : {}) };
  const categoryAxis = (categories: string[]) => ({
    type: 'category' as const,
    data: categories,
    axisLabel: { fontSize: mini ? 8 : 12 },
  });
  const valueAxis = () => ({
    type: 'value' as const,
    axisLabel: { fontSize: mini ? 8 : 12 },
    splitLine: { lineStyle: { color: 'rgba(0,0,0,0.08)' } },
  });
  const grid = mini
    ? { left: 30, right: 6, top: 8, bottom: 16 }
    : { left: 64, right: 28, top: 48, bottom: 40 };

  // 元素开关：标题开启时在顶部显示标题文字
  const titleOption = el.chartTitle
    ? { title: { text: meta?.title || '图表标题', left: 'center', top: 6, textStyle: { fontSize: mini ? 9 : 14, color: '#1f2329' } } }
    : {};

  // 雷达图：指标取类别，每列一个多边形
  if (typeId === 'radar') {
    const { categories, series } = extractCategoryData(matrix);
    const maxByRow = Math.max(...series.flatMap((s) => s.data), 1);
    return {
      ...base,
      ...titleOption,
      legend: { show: !mini && el.legend, bottom: mini ? undefined : 4, textStyle },
      radar: {
        indicator: categories.map((name) => ({ name, max: Math.ceil(maxByRow * 1.2) })),
        radius: mini ? '62%' : '66%',
        axisName: { fontSize: mini ? 8 : 12 },
      },
      series: [
        {
          type: 'radar',
          data: series.map((s) => ({ value: s.data, name: s.name })),
          areaStyle: { opacity: 0.15 },
        },
      ],
    };
  }

  // 饼图 / 圆环图
  if (typeId === 'pie' || typeId === 'pie-doughnut') {
    const data = extractPieData(matrix);
    return {
      ...base,
      ...titleOption,
      legend: { show: !mini && el.legend, orient: 'vertical', right: mini ? 0 : 10, top: 'middle', textStyle },
      series: [
        {
          type: 'pie',
          radius: typeId === 'pie-doughnut' ? ['30%', '64%'] : '62%',
          center: mini ? ['50%', '50%'] : ['38%', '52%'],
          data,
          label: { show: el.dataLabel && !mini, formatter: '{b}: {d}%' },
        },
      ],
    };
  }

  // 散点图
  if (typeId === 'scatter') {
    const { series } = extractScatterData(matrix);
    return {
      ...base,
      ...titleOption,
      grid,
      legend: { show: !mini && el.legend, top: mini ? undefined : 8, textStyle },
      xAxis: { ...valueAxis(), axisLabel: { fontSize: mini ? 8 : 12 }, splitLine: { show: el.gridLine } },
      yAxis: { ...valueAxis(), splitLine: { show: el.gridLine } },
      series: series.map((s) => ({ type: 'scatter', name: s.name, data: s.data, symbolSize: mini ? 4 : 12 })),
    };
  }

  // 其余：柱状 / 折线 / 条形 / 面积（按 variant 变体处理）
  const { categories, series } = extractCategoryData(matrix);
  const stacked = typeId === 'bar-stack' || typeId === 'line-stack' || typeId === 'area-stack' || typeId === 'bar-horizontal-stack';
  const percent = typeId === 'bar-percent';
  const horizontal = typeId === 'bar-horizontal' || typeId === 'bar-horizontal-stack';

  let chartSeries: Record<string, unknown>[] = [];
  if (typeId.startsWith('bar')) {
    chartSeries = series.map((s) => ({
      type: 'bar',
      name: s.name,
      data: s.data,
      barMaxWidth: mini ? 8 : 48,
      ...(stacked || percent ? { stack: 'total' } : {}),
    }));
  } else if (typeId.startsWith('area')) {
    chartSeries = series.map((s) => ({
      type: 'line',
      name: s.name,
      data: s.data,
      areaStyle: {},
      ...(stacked ? { stack: 'total' } : {}),
      showSymbol: false,
    }));
  } else {
    const withSymbol = typeId === 'line-marker';
    chartSeries = series.map((s) => ({
      type: 'line',
      name: s.name,
      data: s.data,
      ...(stacked ? { stack: 'total' } : {}),
      ...(withSymbol ? { symbol: 'circle', symbolSize: mini ? 3 : 8, label: { show: !mini } } : {}),
    }));
  }

  if (percent) {
    // 百分比堆积：每个类别内部归一化为百分比
    const totals = categories.map((_, r) => series.reduce((sum, s) => sum + (s.data[r] ?? 0), 0));
    chartSeries = chartSeries.map((s) => ({
      ...s,
      data: (s.data as number[]).map((v, r) => (totals[r] ? +((v / totals[r]) * 100).toFixed(2) : 0)),
    }));
  }

  if (el.dataLabel && !mini) {
    chartSeries = chartSeries.map((s) => ({ ...s, label: { show: true } }));
  }

  if (el.trendline && (typeId.startsWith('bar') || typeId.startsWith('line') || typeId.startsWith('area'))) {
    // 线性趋势线：每个系列一条虚线（不进图例）
    chartSeries.push(
      ...series.map((s) => ({
        type: 'line',
        name: `趋势(${s.name})`,
        data: trendData(s.data),
        showSymbol: false,
        lineStyle: { type: 'dashed' as const, opacity: 0.7 },
        ...(typeId.startsWith('area') ? { areaStyle: { opacity: 0 } } : {}),
      })),
    );
  }

  const categoryAxisOpt = {
    ...categoryAxis(categories),
    axisLine: { show: el.axis },
    axisTick: { show: el.axis },
    axisLabel: { ...categoryAxis(categories).axisLabel, show: el.axis },
  };
  const valueAxisOpt = {
    ...valueAxis(),
    axisLabel: { ...valueAxis().axisLabel, show: el.axis },
    splitLine: { ...valueAxis().splitLine, show: el.gridLine },
    ...(el.axisTitle ? { name: '数值' } : {}),
  };
  return {
    ...base,
    ...titleOption,
    grid,
    legend: { show: !mini && el.legend, top: mini ? undefined : 8, textStyle, data: series.map((s) => s.name) },
    ...(horizontal
      ? { xAxis: valueAxisOpt, yAxis: { ...categoryAxisOpt, data: categories } }
      : { xAxis: categoryAxisOpt, yAxis: valueAxisOpt }),
    series: chartSeries,
  };
}
