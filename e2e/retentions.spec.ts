// 条件格式 / 数据验证回归：工具栏入口打开 Univer 面板；xlsx 导入的规则落位后
// 随快照持久化（刷新后仍在）。规则渲染依赖单元格类型补全，见 normalizeSnapshotCellTypes。
import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { menuClick, waitEditor } from './helpers';

async function buildRulesFixture(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('规则页');
  ws.getCell('A1').value = 1;
  ws.getCell('A2').value = 2;
  ws.getCell('A3').value = 3;
  ws.getCell('B1').value = '甲';
  ws.getCell('B2').value = '乙';
  ws.getCell('B3').value = '丙';
  const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());

  const stylesPath = 'xl/styles.xml';
  const styles = (await zip.file(stylesPath)!.async('string')).replace(
    /<dxfs[^>]*\/>/,
    '<dxfs count="1"><dxf><font><b/><color rgb="FFFF0000"/></font>' +
      '<fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf></dxfs>',
  );
  zip.file(stylesPath, styles);

  const sheetPath = 'xl/worksheets/sheet1.xml';
  const sheet = (await zip.file(sheetPath)!.async('string')).replace(
    '</worksheet>',
    '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="cellIs" operator="greaterThanOrEqual" priority="1" dxfId="0"><formula>2</formula></cfRule>' +
      '</conditionalFormatting>' +
      '<dataValidations count="1">' +
      '<dataValidation type="list" allowBlank="1" sqref="B1:B3"><formula1>"甲,乙,丙"</formula1></dataValidation>' +
      '</dataValidations></worksheet>',
  );
  zip.file(sheetPath, sheet);
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** 读 IndexedDB 中最近更新的文档快照（导入后规则随文档保存） */
async function latestSnapshot(page: Page): Promise<Record<string, unknown>> {
  const docs = await page.evaluate(async () => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const req = indexedDB.open('doculuma-excel');
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    const store = db.transaction('documents', 'readonly').objectStore('documents');
    return new Promise<{ title: string; updatedAt?: number; snapshot?: unknown }[]>((resolve, reject) => {
      const req = store.getAll();
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  });
  const latest = [...docs].sort((a, b) => (b.updatedAt ?? 0) - (a.updatedAt ?? 0))[0];
  return (latest?.snapshot ?? {}) as Record<string, unknown>;
}

function resourceNames(snapshot: Record<string, unknown>): string[] {
  const resources = snapshot.resources as { name: string }[] | undefined;
  return resources?.map((r) => r.name) ?? [];
}

test.describe('条件格式 / 数据验证', () => {
  test('工具栏入口打开新建面板', async ({ page }) => {
    await waitEditor(page);
    const toolbar = page.locator('.myf-toolbar:not(.myf-measure)');

    await toolbar.getByTitle('条件格式').click();
    await page.getByText('新建规则', { exact: true }).click();
    await expect(page.getByText('应用范围', { exact: true })).toBeVisible({ timeout: 5000 });
    await page.keyboard.press('Escape');
    await page.waitForTimeout(300);

    await toolbar.getByTitle('数据验证').click();
    await page.getByText('新建验证规则', { exact: true }).click();
    await expect(page.getByText('条件类型', { exact: true })).toBeVisible({ timeout: 5000 });
  });

  test('xlsx 导入的条件格式/数据验证落位并持久化', async ({ page }) => {
    await waitEditor(page);
    await menuClick(page, '文件', '导入');
    const dialog = page.getByRole('dialog');
    await dialog.waitFor({ timeout: 5000 });
    await dialog.locator('input[type="file"]').setInputFiles({
      name: '规则夹具.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: await buildRulesFixture(),
    });
    await dialog.getByRole('button', { name: /确认导入/ }).click();
    await expect(page.getByText('导入成功', { exact: true })).toBeVisible({ timeout: 20000 });
    // 规则入队落位 + 立即 saveNow
    await page.waitForTimeout(4000);

    const before = resourceNames(await latestSnapshot(page));
    expect(before).toContain('SHEET_CONDITIONAL_FORMATTING_PLUGIN');
    expect(before).toContain('SHEET_DATA_VALIDATION_PLUGIN');

    await page.reload();
    await page.locator('#univer-container canvas:visible').first().waitFor({ timeout: 30000 });
    await page.getByTitle('当前单元格').waitFor({ timeout: 15000 });
    await page.waitForTimeout(1500);
    const after = resourceNames(await latestSnapshot(page));
    expect(after).toContain('SHEET_CONDITIONAL_FORMATTING_PLUGIN');
    expect(after).toContain('SHEET_DATA_VALIDATION_PLUGIN');
  });
});
