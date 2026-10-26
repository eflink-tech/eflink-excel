// xlsx 导入端到端：共享公式/隐藏 sheet 的单元格保真 + 图表落位为浮动图片并随快照持久化。
// fixture 在测试内现场构造（exceljs 造数据 → jszip 注入共享公式与 chart/drawing 部件），不落仓库文件。
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { clickCellA1, menuClick, waitEditor, nameBox, typeIntoFormulaBar } from './helpers';

const CHART_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>e2e!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>频数</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>e2e!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>甲</c:v></c:pt><c:pt idx="1"><c:v>乙</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>e2e!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>50</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart><c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;

async function buildFixtureXlsx(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('e2e');
  ws.getCell('A1').value = '类别';
  ws.getCell('B1').value = '频数';
  ws.getCell('A2').value = '甲';
  ws.getCell('A3').value = '乙';
  ws.getCell('B2').value = 50;
  ws.getCell('B3').value = 30;
  ws.getCell('C2').value = { formula: 'B2*2', result: 100 };
  ws.getCell('C3').value = { formula: 'B3*2', result: 60 };
  const hidden = wb.addWorksheet('隐藏页');
  hidden.getCell('A1').value = 'x';
  hidden.state = 'hidden';
  const buffer = await wb.xlsx.writeBuffer();

  const zip = await JSZip.loadAsync(buffer);
  // C 列改写为共享公式（主 C2 + 依赖 C3），与 Excel/WPS 生产文件同构
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  const rewritten = sheetXml
    .replace(/<c r="C2"[^>]*>.*?<\/c>/, '<c r="C2"><f t="shared" ref="C2:C3" si="0">B2*2</f><v>100</v></c>')
    .replace(/<c r="C3"[^>]*>.*?<\/c>/, '<c r="C3"><f t="shared" si="0"/><v>60</v></c>');
  if (rewritten === sheetXml) throw new Error('共享公式改写未命中');
  zip.file('xl/worksheets/sheet1.xml', rewritten);

  // 注入 drawing 部件：一张锚定 F2 的簇状柱状图
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', ct
    .replace('</Types>', '<Default Extension="xml" PartName="/xl/drawings/drawing1.xml"/><Default Extension="xml" PartName="/xl/charts/chart1.xml"/></Types>'));
  const sheetRelsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  const existing = await zip.file(sheetRelsPath)?.async('string');
  const drawingRel = '<Relationship Id="rIdDr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>';
  zip.file(sheetRelsPath, existing ? existing.replace('</Relationships>', drawingRel + '</Relationships>') : `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawingRel}</Relationships>`);
  const sheetPathXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file('xl/worksheets/sheet1.xml', sheetPathXml.replace(/(<worksheet[^>]*>)/, '$1<drawing r:id="rIdDr1"/>'));
  zip.file('xl/drawings/drawing1.xml', `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>14</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdCh1"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>`);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCh1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`);
  zip.file('xl/charts/chart1.xml', CHART_XML);
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** 走菜单导入 fixture，返回导入完成后的页面（图表落位等待由调用方控制） */
async function importFixture(page: import('@playwright/test').Page): Promise<void> {
  const fixture = await buildFixtureXlsx();
  await menuClick(page, '文件');
  const fileChooserPromise = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: '导入 Excel(.xlsx)' }).click();
  const chooser = await fileChooserPromise;
  await chooser.setFiles({ name: '导入夹具.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: fixture });
  const confirm = page.getByRole('button', { name: '确定' });
  await expect(confirm).toBeVisible({ timeout: 5000 });
  await confirm.click();
  await expect(page.getByText('导入成功（含 1 张图表）')).toBeVisible({ timeout: 15000 });
}

/** 导出 efexcel 并解析 payload（导出动作内部先 saveNow，必为最新快照） */
async function exportPayload(page: import('@playwright/test').Page): Promise<{ snapshot: Record<string, unknown> }> {
  const downloadPromise = page.waitForEvent('download', { timeout: 20000 });
  await menuClick(page, '文件', '导出数据(.efx.json)');
  const download = await downloadPromise;
  return JSON.parse((await import('node:fs')).readFileSync(await download.path(), 'utf8')) as never;
}

test.describe('导入 Excel(.xlsx)', () => {
  test('共享公式/隐藏 sheet 保真，图表落位为浮动图片', async ({ page }) => {
    await waitEditor(page);
    await importFixture(page);

    // 图表插入为浮动图片并随快照持久化：等落位与自动保存后导出 efexcel 检查
    await page.waitForTimeout(3500);
    const payload = await exportPayload(page);
    const snap = payload.snapshot as {
      sheets: Record<string, { name: string; hidden?: number; cellData: Record<number, Record<number, { f?: string; v?: unknown }>> }>;
      [key: string]: unknown;
    };
    const sheetIds = Object.keys(snap.sheets);
    expect(sheetIds.length).toBe(2);
    const main = Object.values(snap.sheets).find((s) => s.name === 'e2e')!;
    // 共享公式恢复：主/依赖单元格各自公式 + 缓存值
    expect(main.cellData[1]?.[2]).toMatchObject({ f: '=B2*2', v: 100 });
    expect(main.cellData[2]?.[2]).toMatchObject({ f: '=B3*2', v: 60 });
    // 隐藏 sheet 随快照持久化
    expect(Object.values(snap.sheets).find((s) => s.name === '隐藏页')?.hidden).toBe(1);
    // 图表图片进入快照（Univer drawings 资源；具体键名随版本，做存在性探测）
    // 图表图片经 Univer resources 插件资源区随快照持久化（刷新后仍在）
    const raw = JSON.stringify(payload);
    expect(Object.keys(snap)).toContain('resources');
    expect(raw.length).toBeGreaterThan(50000);
  });

  test('数据联动：编辑图表源单元格后自动重渲染', async ({ page }) => {
    await waitEditor(page);
    await importFixture(page);
    await page.waitForTimeout(3500);

    const before = await exportPayload(page);
    const imagesBefore = JSON.stringify(before).match(/data:image\/png;base64,[A-Za-z0-9+/=]+/g) ?? [];
    expect(imagesBefore.length).toBeGreaterThan(0);

    // 编辑图表源数据 B2（频数 50 → 80）：clickCellA1 建立键盘路由后方向键导航
    await clickCellA1(page);
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowDown');
    await expect(nameBox(page)).toHaveValue('B2');
    await typeIntoFormulaBar(page, '80');
    // 联动防抖 700ms + ECharts 渲染 + 自动保存余量
    await page.waitForTimeout(4500);

    const after = await exportPayload(page);
    const imagesAfter = JSON.stringify(after).match(/data:image\/png;base64,[A-Za-z0-9+/=]+/g) ?? [];
    // 单元格值已变
    const main = Object.values((after.snapshot as { sheets: Record<string, { name: string; cellData: Record<number, Record<number, { v?: unknown }>> }> }).sheets).find((sh) => sh.name === 'e2e')!;
    expect(main.cellData[1]?.[1]?.v).toBe(80);
    // 图表图片内容随数据变化（同数据同管线为确定性渲染，串相同即联动失效）
    expect(imagesAfter.join()).not.toBe(imagesBefore.join());
  });
});
