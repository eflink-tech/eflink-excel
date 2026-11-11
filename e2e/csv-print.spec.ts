// CSV 导入导出 + 打印 / PDF 端到端：导入弹窗预览、确认后落库；导出带 BOM/CRLF 并正确转义；
// 打印弹窗预览当前工作表，空表禁用打印；.xls 旧格式明确拒绝。
import { expect, test, type Page } from '@playwright/test';
import { menuClick, typeIntoFormulaBar, waitEditor } from './helpers';

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

test.describe('CSV 导入 / 导出', () => {
  test('CSV 导入先预览，确认后按文本写入并保留前导零', async ({ page }) => {
    await waitEditor(page);
    await menuClick(page, '文件', '导入');

    const csv = '名称,数量\r\n"带,逗号",007\r\n"他说""嗨""",3\r\n';
    await page.locator('input[accept=".xlsx,.xlsm,.csv,.json"]').setInputFiles({
      name: '清单.csv',
      mimeType: 'text/csv',
      buffer: Buffer.from(csv, 'utf8'),
    });

    await expect(page.getByText('导入预览')).toBeVisible({ timeout: 5000 });
    const dialog = page.getByRole('dialog', { name: '导入预览' });
    await expect(dialog.getByText('带,逗号')).toBeVisible();
    await expect(dialog.getByText('他说"嗨"')).toBeVisible();

    await page.getByRole('button', { name: '确认导入（可编辑）' }).click();
    await expect(page.getByText('导入成功', { exact: true })).toBeVisible({ timeout: 15000 });

    const snapshot = await latestSnapshot(page);
    const sheets = snapshot.sheets as Record<string, { name: string; cellData: Record<number, Record<number, { v?: unknown; t?: number }>> }>;
    const sheet = sheets[Object.keys(sheets)[0]];
    expect(sheet.name).toBe('清单');
    expect(sheet.cellData[0][1].v).toBe('数量');
    expect(sheet.cellData[1][1].v).toBe('007');
    expect(sheet.cellData[1][1].t).toBe(1);
    expect(sheet.cellData[1][0].v).toBe('带,逗号');
    expect(sheet.cellData[2][0].v).toBe('他说"嗨"');
  });

  test('导出 CSV 带 BOM 与 CRLF，按 RFC4180 转义', async ({ page }) => {
    await waitEditor(page);
    await typeIntoFormulaBar(page, 'a,b');
    await page.keyboard.press('ArrowDown');
    await typeIntoFormulaBar(page, 'he said "hi"');

    const downloadPromise = page.waitForEvent('download', { timeout: 15000 });
    await menuClick(page, '文件', '导出');
    await page.getByRole('button', { name: /CSV \(\.csv\)/ }).click();
    const download = await downloadPromise;
    const fs = await import('node:fs');
    const text = fs.readFileSync(await download.path(), 'utf8');

    expect(text.charCodeAt(0)).toBe(0xfeff);
    expect(text).toContain('"a,b"');
    expect(text).toContain('"he said ""hi"""');
    expect(text).toContain('\r\n');
    await expect(page.getByText('已导出 CSV')).toBeVisible({ timeout: 5000 });
  });

  test('.xls 旧格式提示另存为 .xlsx', async ({ page }) => {
    await waitEditor(page);
    await menuClick(page, '文件', '导入');
    await page.locator('input[accept=".xlsx,.xlsm,.csv,.json"]').setInputFiles({
      name: '旧表.xls',
      mimeType: 'application/vnd.ms-excel',
      buffer: Buffer.from('not really xls'),
    });
    await expect(
      page.getByText('不支持 .xls 旧格式，请先在 Excel / WPS 中另存为 .xlsx 后再导入'),
    ).toBeVisible({ timeout: 5000 });
  });
});

test.describe('打印 / PDF', () => {
  test('空表禁用打印，有内容后可预览并切换纸张方向', async ({ page }) => {
    await waitEditor(page);
    await menuClick(page, '文件', '打印 / PDF');
    const dialog = page.getByRole('dialog', { name: '打印 / PDF' });
    await expect(dialog.getByText('当前工作表没有可打印内容')).toBeVisible({ timeout: 5000 });
    await expect(dialog.getByRole('button', { name: '打印', exact: true })).toBeDisabled();
    await dialog.getByRole('button', { name: '取消', exact: true }).click();

    await typeIntoFormulaBar(page, '打印测试');
    await menuClick(page, '文件', '打印 / PDF');
    await expect(dialog).toBeVisible({ timeout: 5000 });
    const frame = page.frameLocator('iframe[title="打印预览"]');
    await expect(frame.getByText('打印测试')).toBeVisible({ timeout: 5000 });

    await dialog.getByLabel('横向').check();
    await expect
      .poll(() => frame.locator('style').evaluate((el) => el.textContent ?? ''))
      .toContain('size: A4 landscape');
    await expect(dialog.getByRole('button', { name: '另存为 PDF' })).toBeEnabled();
  });

  test('导出弹窗的 PDF 选项进入打印预览', async ({ page }) => {
    await waitEditor(page);
    await typeIntoFormulaBar(page, 'PDF');
    await menuClick(page, '文件', '导出');
    await page.getByRole('button', { name: /PDF\(\.pdf\)/ }).click();
    const dialog = page.getByRole('dialog', { name: '打印 / PDF' });
    await expect(dialog).toBeVisible({ timeout: 5000 });
    await expect(page.frameLocator('iframe[title="打印预览"]').getByText('PDF')).toBeVisible({ timeout: 5000 });
  });
});
