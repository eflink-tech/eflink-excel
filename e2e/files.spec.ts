import fs from 'node:fs';
import { expect, test } from '@playwright/test';
import { clickCellA1, contentInput, menuClick, nameBox, waitEditor } from './helpers';

function firstSheetOf(json: string): { cellData: Record<number, Record<number, { v?: unknown; f?: string }>> } {
  const doc = JSON.parse(json);
  expect(doc.title).toBeTruthy();
  expect(doc.snapshot.sheetOrder.length).toBeGreaterThan(0);
  return doc.snapshot.sheets[doc.snapshot.sheetOrder[0]];
}

test.describe('文件导入导出', () => {
  test('导出 efexcel：标题与单元格内容写入 JSON 文件', async ({ page }) => {
    await waitEditor(page);

    await contentInput(page).fill('备份内容');
    await contentInput(page).press('Enter');

    const downloadPromise = page.waitForEvent('download', { timeout: 15000 });
    await menuClick(page, '文件', '导出表格(.efexcel)');
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toMatch(/\.efexcel$/);

    const sheet = firstSheetOf(fs.readFileSync(await download.path(), 'utf8'));
    expect(sheet.cellData[0]?.[0]?.v).toBe('备份内容');
    await expect(page.getByText('已导出 efexcel')).toBeVisible({ timeout: 5000 });
  });

  test('导出 efexcel：协议携带 source 来源信息', async ({ page }) => {
    await waitEditor(page);

    const downloadPromise = page.waitForEvent('download', { timeout: 15000 });
    await menuClick(page, '文件', '导出表格(.efexcel)');
    const download = await downloadPromise;

    const payload = JSON.parse(fs.readFileSync(await download.path(), 'utf8'));
    expect(payload.source).toEqual({ app: '易飞表格', url: 'https://eflink.tech/office/excel' });
  });

  test('导出 efexcel：公式与计算结果随快照导出', async ({ page }) => {
    await waitEditor(page);

    await contentInput(page).fill('=1+2');
    await contentInput(page).press('Enter');
    // Univer 的计算结果异步写回快照（cellData.v），等待后再导出才能携带缓存结果
    await page.waitForTimeout(2000);

    const downloadPromise = page.waitForEvent('download', { timeout: 15000 });
    await menuClick(page, '文件', '导出表格(.efexcel)');
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toMatch(/\.efexcel$/);

    const sheet = firstSheetOf(fs.readFileSync(await download.path(), 'utf8'));
    const cell = sheet.cellData[0]?.[0];
    expect(cell?.f).toBeTruthy();
    expect(cell?.v).toBe(3);
  });

  test('导出图片：html2canvas 截图下载 png', async ({ page }) => {
    await waitEditor(page);

    const downloadPromise = page.waitForEvent('download', { timeout: 30000 });
    await menuClick(page, '文件', '导出图片');
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toBe('未命名表格.png');
    await expect(page.getByText('已导出图片')).toBeVisible({ timeout: 5000 });
  });

  test('导入 efexcel：新建文档打开并回显导入内容', async ({ page }) => {
    await waitEditor(page);

    const doc = {
      id: 'doc-src',
      title: '导入报表',
      createdAt: Date.now(),
      updatedAt: Date.now(),
      snapshot: {
        id: 'wb-src',
        name: '导入报表',
        sheetOrder: ['sheet-01'],
        sheets: {
          'sheet-01': {
            id: 'sheet-01',
            name: 'Sheet1',
            rowCount: 100,
            columnCount: 20,
            cellData: { 0: { 0: { v: '导入标题' }, 1: { v: 42 } } },
          },
        },
      },
    };
    await page.setInputFiles('input[type="file"]', {
      name: '导入报表.efexcel',
      mimeType: 'application/json',
      buffer: Buffer.from(JSON.stringify(doc)),
    });

    await expect(page.getByText('导入成功')).toBeVisible({ timeout: 10000 });
    await expect(nameBox(page)).toHaveValue('A1', { timeout: 15000 });
    await expect(contentInput(page)).toHaveValue('导入标题');
  });

  test('导入损坏文件提示导入失败且不进入新文档', async ({ page }) => {
    await waitEditor(page);

    await page.setInputFiles('input[type="file"]', {
      name: '坏文件.efexcel',
      mimeType: 'application/json',
      buffer: Buffer.from('这不是一个 efexcel 文件'),
    });

    await expect(page.getByText('导入失败：无法解析该文件')).toBeVisible({ timeout: 10000 });
  });

  test('导入缺少快照结构的 JSON 提示导入失败', async ({ page }) => {
    await waitEditor(page);
    await clickCellA1(page);

    await page.setInputFiles('input[type="file"]', {
      name: '空文档.efexcel',
      mimeType: 'application/json',
      buffer: Buffer.from(JSON.stringify({ title: '空文档' })),
    });

    await expect(page.getByText('导入失败：无法解析该文件')).toBeVisible({ timeout: 10000 });
  });
});
