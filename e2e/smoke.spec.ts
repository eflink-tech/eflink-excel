import { expect, test } from '@playwright/test';

test('打开应用直接进入编辑器（企微式布局）', async ({ page }) => {
  await page.goto('/');
  // 无列表页：Univer 画布直接挂载
  await expect(page.locator('#univer-container canvas:visible').first()).toBeVisible({
    timeout: 30000,
  });
  // 顶部：office-kit 统一顶栏（品牌名 + 横向主菜单）
  await expect(page.locator('.efk-brand-name')).toHaveText('DocuLuma Sheets');
  await expect(page.getByRole('menuitem', { name: '文件' })).toBeVisible();
  // 工具栏：myf-sheet-demo 同款单行工具栏（插入面板 + 数字格式面板）
  await expect(page.getByRole('button', { name: '插入' })).toBeVisible();
  await expect(page.getByRole('button', { name: '数字格式' })).toBeVisible();
  await page.getByRole('button', { name: '插入' }).click();
  await expect(page.getByText('在上方插入行')).toBeVisible();

  // 主菜单「文件」含导出项（office-kit 菜单按钮显式 role="menuitem"）
  await page.getByRole('menuitem', { name: '文件' }).click();
  await expect(page.getByRole('menuitem', { name: '导出数据 (.efx.json)' })).toBeVisible();
  await page.keyboard.press('Escape');
});
