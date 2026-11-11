// 单元格内富文本：F2 编辑 + 选中部分文字 + 工具栏加粗，只改选中文字并随快照持久化。
import { expect, test, type Page } from '@playwright/test';
import { nameBox, typeIntoFormulaBar, waitEditor } from './helpers';

const toolbar = (page: Page) => page.locator('.myf-toolbar:not(.myf-measure)');

test('编辑单元格时工具栏加粗只作用于选中文字', async ({ page }) => {
  await waitEditor(page);
  await typeIntoFormulaBar(page, 'ABCDEF');

  await page.keyboard.press('F2');
  await expect
    .poll(() => page.evaluate(() => document.activeElement?.getAttribute('data-u-comp')))
    .toBe('editor');
  await page.keyboard.press('Shift+ArrowLeft');
  await page.keyboard.press('Shift+ArrowLeft');
  await toolbar(page).locator('a[title="加粗"]').first().click();
  await page.waitForTimeout(300);
  // 工具栏点击不应抢走编辑焦点，否则命令会落到整格
  await expect
    .poll(() => page.evaluate(() => document.activeElement?.getAttribute('data-u-comp')))
    .toBe('editor');
  await page.keyboard.press('Enter');
  await expect(nameBox(page)).toHaveValue('A2');

  await page.keyboard.press('ControlOrMeta+s');
  await page.waitForTimeout(1200);

  const cell = await readA1(page);
  expect(cell.s).toBeUndefined();
  expect(cell.p?.body?.dataStream).toBe('ABCDEF\r\n');
  expect(cell.p?.body?.textRuns).toEqual([{ st: 4, ed: 6, ts: { bl: 1 } }]);
});

test('富文本刷新后仍保留', async ({ page }) => {
  await waitEditor(page);
  await typeIntoFormulaBar(page, 'HELLO');

  await page.keyboard.press('F2');
  await expect
    .poll(() => page.evaluate(() => document.activeElement?.getAttribute('data-u-comp')))
    .toBe('editor');
  for (let i = 0; i < 3; i++) await page.keyboard.press('Shift+ArrowLeft');
  await toolbar(page).locator('a[title="倾斜"]').first().click();
  await page.waitForTimeout(300);
  await page.keyboard.press('Enter');
  await page.keyboard.press('ControlOrMeta+s');
  await page.waitForTimeout(1200);

  await page.reload();
  await expect(page.locator('#univer-container canvas:visible').first()).toBeVisible({ timeout: 30000 });
  await page.waitForTimeout(1200);
  const cell = await readA1(page);
  expect(cell.p?.body?.textRuns).toEqual([{ st: 2, ed: 5, ts: { it: 1 } }]);
});

interface StoredCell {
  v?: unknown;
  s?: unknown;
  p?: { body?: { dataStream?: string; textRuns?: unknown[] } };
}

async function readA1(page: Page): Promise<StoredCell> {
  const snapshot = await page.evaluate(async () => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const req = indexedDB.open('doculuma-excel');
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    const store = db.transaction('documents', 'readonly').objectStore('documents');
    const docs = await new Promise<{ updatedAt?: number; snapshot?: Record<string, unknown> }[]>((resolve, reject) => {
      const req = store.getAll();
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    const latest = [...docs].sort((a, b) => (b.updatedAt ?? 0) - (a.updatedAt ?? 0))[0];
    return latest?.snapshot ?? {};
  });
  const sheets = snapshot.sheets as Record<string, { cellData: Record<number, Record<number, StoredCell>> }>;
  return sheets[Object.keys(sheets)[0]].cellData[0]?.[0] ?? {};
}
