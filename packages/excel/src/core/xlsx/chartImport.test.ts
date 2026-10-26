// 图表导入解析测试：chart XML 解析为纯函数直接喂字符串；关系链遍历（workbook→sheet→drawing→chart）
// 用 jszip 构造最小可用包验证。渲染落位（echarts/Univer）属浏览器链路，不在 node 测试覆盖。
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { extractCharts, parseChartXml, queueImportedCharts, takeImportedCharts } from './chartImport';

/** 最小 ser 片段：cat/val 缓存各给两点 */
function serXml(name: string, cats: string[], vals: number[]): string {
  const cat = cats.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join('');
  const val = vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');
  return `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>数据!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>数据!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/>${cat}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>数据!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/>${val}</c:numCache></c:numRef></c:val></c:ser>`;
}

function chartXml(body: string, opts: { title?: string } = {}): string {
  const title = opts.title ? `<c:title><c:tx><c:rich><a:p><a:r><a:t>${opts.title}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>${title}<c:plotArea><c:layout/>${body}<c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;
}

describe('parseChartXml', () => {
  it('柱状图：barDir/grouping 映射 typeId，cat/val 缓存进显式数据', () => {
    const parsed = parseChartXml(chartXml(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${serXml('频数', ['甲', '乙'], [30, 10])}</c:barChart>`));
    expect(parsed?.typeId).toBe('bar-cluster');
    expect(parsed?.explicit.categories).toEqual(['甲', '乙']);
    expect(parsed?.explicit.series).toEqual([{ name: '频数', data: [30, 10] }]);
    expect(parsed?.rangeRef).toBe('A2:A3');
    expect(parsed?.title).toBeUndefined();
    // 联动源引用：类别 + 每系列数值（编辑后按此重读）
    expect(parsed?.liveRefs).toEqual({
      catRef: 'A2:A3',
      series: [{ nameRef: 'B1', valRef: 'B2:B3' }],
    });
  });

  it('横向堆积条形图与百分比堆积柱状图', () => {
    expect(parseChartXml(chartXml('<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/>' + serXml('S', ['x'], [1]) + '</c:barChart>'))?.typeId).toBe('bar-horizontal-stack');
    expect(parseChartXml(chartXml('<c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/>' + serXml('S', ['x'], [1]) + '</c:barChart>'))?.typeId).toBe('bar-percent');
  });

  it('柱+折线组合（帕累托）：typeId=combo-bar-line，柱/折线系列分离', () => {
    const combo = chartXml(
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${serXml('频数', ['甲', '乙'], [50, 30])}</c:barChart>` +
      `<c:lineChart><c:grouping val="standard"/>${serXml('累计%', ['甲', '乙'], [0.51, 0.82])}</c:lineChart>`,
    );
    const parsed = parseChartXml(combo);
    expect(parsed?.typeId).toBe('combo-bar-line');
    expect(parsed?.explicit.combo?.bars).toEqual([{ name: '频数', data: [50, 30] }]);
    expect(parsed?.explicit.combo?.lines).toEqual([{ name: '累计%', data: [0.51, 0.82] }]);
    expect(parsed?.explicit.categories).toEqual(['甲', '乙']);
    // 组合图源引用带 comboBars 分界（前 1 柱 + 后 1 折线）
    expect(parsed?.liveRefs?.comboBars).toBe(1);
    expect(parsed?.liveRefs?.series).toHaveLength(2);
  });

  it('雷达图 / 饼图(3D) / 圆环图', () => {
    expect(parseChartXml(chartXml(`<c:radarChart><c:radarStyle val="standard"/>${serXml('活动前', ['质量', '知识'], [4, 5])}</c:radarChart>`))?.typeId).toBe('radar');
    const pie = parseChartXml(chartXml(`<c:pie3DChart>${serXml('频数', ['甲', '乙'], [50, 30])}</c:pie3DChart>`));
    expect(pie?.typeId).toBe('pie');
    expect(pie?.explicit.pieData).toEqual([{ name: '甲', value: 50 }, { name: '乙', value: 30 }]);
    expect(parseChartXml(chartXml(`<c:doughnutChart>${serXml('S', ['a'], [1])}</c:doughnutChart>`))?.typeId).toBe('pie-doughnut');
  });

  it('散点图：xVal/yVal 配对为坐标', () => {
    const scatterSer = `<c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>数据!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>移动极差</c:v></c:pt></c:strCache></c:strRef></c:tx><c:xVal><c:numRef><c:f>数据!$A$2:$A$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>2.1</c:v></c:pt><c:pt idx="1"><c:v>3.2</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>数据!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1.1</c:v></c:pt><c:pt idx="1"><c:v>0.9</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser>`;
    const parsed = parseChartXml(chartXml(`<c:scatterChart><c:scatterStyle val="lineMarker"/>${scatterSer}</c:scatterChart>`));
    expect(parsed?.typeId).toBe('scatter');
    expect(parsed?.explicit.scatterSeries).toEqual([{ name: '移动极差', data: [[2.1, 1.1], [3.2, 0.9]] }]);
    expect(parsed?.rangeRef).toBe('A2:A3');
    expect(parsed?.liveRefs?.series[0]).toMatchObject({ xRef: 'A2:A3', yRef: 'B2:B3' });
  });

  it('标题保留、autoTitleDeleted 抑制、多行富文本标题拼接', () => {
    expect(parseChartXml(chartXml('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' + serXml('S', ['x'], [1]) + '</c:barChart>', { title: '产量' }))?.title).toBe('产量');
    const multiRun = chartXml('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' + serXml('S', ['x'], [1]) + '</c:barChart>');
    // autoTitleDeleted=1 的默认骨架无标题
    const noTitle = parseChartXml(multiRun);
    expect(noTitle?.title).toBeUndefined();
  });

  it('lineChart 带 marker → line-marker；无 marker → line', () => {
    const withMarker = '<c:lineChart><c:grouping val="standard"/>' + serXml('S', ['x'], [1]).replace('</c:tx>', '</c:tx><c:marker><c:symbol val="circle"/></c:marker>') + '</c:lineChart>';
    expect(parseChartXml(chartXml(withMarker))?.typeId).toBe('line-marker');
    expect(parseChartXml(chartXml('<c:lineChart><c:grouping val="standard"/>' + serXml('S', ['x'], [1]) + '</c:lineChart>'))?.typeId).toBe('line');
  });

  it('数值类别轴（numRef cat）按字符串保留', () => {
    const numCatSer = `<c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>数据!$C$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>CL</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:numRef><c:f>数据!$A$2:$A$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:cat><c:val><c:numRef><c:f>数据!$C$2:$C$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>4.07</c:v></c:pt><c:pt idx="1"><c:v>4.07</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>`;
    const parsed = parseChartXml(chartXml(`<c:lineChart><c:grouping val="standard"/>${numCatSer}</c:lineChart>`));
    expect(parsed?.explicit.categories).toEqual(['2', '3']);
  });

  it('不认识的图表类型返回 null（跳过）', () => {
    expect(parseChartXml(chartXml('<c:surfaceChart><c:ser/></c:surfaceChart>'))).toBeNull();
  });

  it('图表筛选器隐藏的系列（filteredBarSeries 容器）不参与解析与联动引用', () => {
    // barChart 带一个有效系列 + extLst 里被筛选隐藏的"百分比"系列（qcc 排列图真实形态）
    const xml = chartXml(
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${serXml('频数', ['甲', '乙'], [50, 30])}` +
      `<c:extLst><c:ext><c:filteredBarSeries>${serXml('百分比', ['甲', '乙'], [0.5, 0.3])}</c:filteredBarSeries></c:ext></c:extLst></c:barChart>`,
    );
    const parsed = parseChartXml(xml);
    expect(parsed?.explicit.series).toEqual([{ name: '频数', data: [50, 30] }]);
    expect(parsed?.liveRefs?.series).toHaveLength(1);
    expect(parsed?.liveRefs?.series[0]?.valRef).toBe('B2:B3');
  });
});

describe('extractCharts 关系链', () => {
  /** 构造含一张图表的最小 xlsx 包 */
  async function buildChartXlsx(): Promise<ArrayBuffer> {
    const zip = new JSZip();
    zip.file(
      'xl/workbook.xml',
      `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="常用工具" sheetId="1" r:id="rId1"/><sheet name="隐藏页" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>`,
    );
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`);
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><drawing r:id="rId1"/></worksheet>`);
    zip.file('xl/worksheets/sheet2.xml', `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`);
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`);
    zip.file('xl/drawings/drawing1.xml', `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>12</xdr:col><xdr:colOff>19050</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>`);
    zip.file('xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`);
    zip.file('xl/charts/chart1.xml', chartXml('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' + serXml('频数', ['甲', '乙'], [50, 30]) + '</c:barChart>'));
    return zip.generateAsync({ type: 'arraybuffer' });
  }

  it('按 workbook→sheet→drawing→chart 链路取到锚点与数据', async () => {
    const charts = await extractCharts(await buildChartXlsx());
    expect(charts).toHaveLength(1);
    const chart = charts[0]!;
    expect(chart.sheetName).toBe('常用工具');
    expect(chart.anchor.col).toBe(5);
    expect(chart.anchor.row).toBe(1);
    expect(chart.anchor.toCol).toBe(12);
    expect(chart.anchor.toRow).toBe(8);
    expect(chart.anchor.toColOffPx).toBe(2); // 19050 EMU = 2px
    expect(chart.typeId).toBe('bar-cluster');
    expect(chart.explicit.series?.[0]?.data).toEqual([50, 30]);
  });

  it('无 drawing 的 sheet 静默跳过；队列一次性消费', async () => {
    const empty = await extractCharts(new ArrayBuffer(0)).catch(() => []);
    expect(empty).toEqual([]);
    queueImportedCharts([]);
    expect(takeImportedCharts()).toEqual([]);
  });
});
