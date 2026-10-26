// xlsx 导入图表的落位：导入动作只产出「锚点 + 数据」，图片插入必须等 Univer 挂载完成。
// fileIO 渲染好 dataUrl 后经 chartImport 的队列传递，UniverContainer 每次挂载后调用
// insertQueuedImportCharts 消费：按锚点插浮动图片、按 twoCellAnchor 区间量尺寸、登记设置面板元数据。
import { getUniverApi } from './controller';
import { setChartMeta, defaultChartElements } from './chartSettings';
import { takeImportedCharts, type ImportedChart } from '../xlsx/chartImport';

// facade 类型从 API 推导，避免依赖各包类型导出路径
type UniverApi = NonNullable<ReturnType<typeof getUniverApi>>;
type WorkbookLike = NonNullable<ReturnType<UniverApi['getActiveWorkbook']>>;
type WorksheetLike = NonNullable<ReturnType<WorkbookLike['getSheetByName']>>;

/** 锚点区间尺寸下限：避免空区间/异常锚点把图片缩成看不见 */
const MIN_WIDTH_PX = 120;
const MIN_HEIGHT_PX = 90;

/** twoCellAnchor 区间 → 像素尺寸（列宽/行高按当前工作表实际值累计，EMU 偏移换算像素） */
function measureAnchor(ws: WorksheetLike, chart: ImportedChart): { width: number; height: number } {
  const a = chart.anchor;
  if (a.extWidthPx && a.extHeightPx) return { width: a.extWidthPx, height: a.extHeightPx };
  const toCol = a.toCol ?? a.col;
  const toRow = a.toRow ?? a.row;
  let width = (a.toColOffPx ?? 0) - a.colOffPx;
  for (let c = a.col; c < toCol; c++) width += ws.getColumnWidth(c);
  let height = (a.toRowOffPx ?? 0) - a.rowOffPx;
  for (let r = a.row; r < toRow; r++) height += ws.getRowHeight(r);
  return { width: Math.max(width, MIN_WIDTH_PX), height: Math.max(height, MIN_HEIGHT_PX) };
}

/** 导入图表 → 设置面板元数据的 matrix 形状（首行系列名、首列类别名；面板降级重渲染用） */
function metaMatrix(chart: ImportedChart): unknown[][] {
  const e = chart.explicit;
  if (e.combo) {
    const { bars, lines } = e.combo;
    const names = [...bars, ...lines].map((s) => s.name);
    const rows = (e.categories ?? []).map((cat, i) => [cat, ...bars.map((b) => b.data[i] ?? 0), ...lines.map((l) => l.data[i] ?? 0)]);
    return [['', ...names], ...rows];
  }
  if (e.scatterSeries) {
    const first = e.scatterSeries[0];
    return [['参数1', first?.name ?? '参数2'], ...(first?.data.map(([x, y]) => [x, y]) ?? [])];
  }
  if (e.pieData) return e.pieData.map((d) => [d.name, d.value]);
  if (e.categories && e.series) {
    return [['', ...e.series.map((s) => s.name)], ...e.categories.map((cat, i) => [cat, ...e.series!.map((s) => s.data[i] ?? 0)])];
  }
  return [];
}

/** 消费导入队列：把渲染好的图表逐张落位。任何一张失败只告警，不影响其余 */
export function insertQueuedImportCharts(): void {
  const charts = takeImportedCharts();
  if (!charts.length) return;
  void (async () => {
    const api = getUniverApi();
    const wb = api?.getActiveWorkbook();
    if (!wb) return;
    for (const chart of charts) {
      try {
        await insertOne(wb, chart);
      } catch (err) {
        console.warn('[chartImport] 图表落位失败，已跳过', chart.sheetName, chart.rangeRef, err);
      }
    }
  })();
}

async function insertOne(wb: WorkbookLike, chart: ImportedChart): Promise<void> {
  if (!chart.dataUrl) return;
  const ws = wb.getSheetByName(chart.sheetName);
  if (!ws) return;
  const { width, height } = measureAnchor(ws, chart);
  const ok = await ws.insertImage(chart.dataUrl, chart.anchor.col, chart.anchor.row);
  if (!ok) return;
  const images = ws.getImages();
  const added = images[images.length - 1];
  if (!added) return;
  await added.setSizeAsync(Math.round(width), Math.round(height));
  setChartMeta({
    drawingId: added.getId(),
    typeId: chart.typeId,
    matrix: metaMatrix(chart),
    elements: { ...defaultChartElements(), chartTitle: Boolean(chart.title) },
    paletteIndex: 0,
    title: chart.title ?? '',
    rangeRef: chart.rangeRef,
    explicit: chart.explicit,
    sheetName: chart.sheetName,
    liveRefs: chart.liveRefs,
    // 初始数据指纹：首次编辑若值未变则跳过重渲染
    dataSignature: JSON.stringify(chart.explicit),
  });
}
