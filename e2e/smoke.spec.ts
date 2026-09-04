import { expect, test } from '@playwright/test';

test('打开应用直接进入编辑器（企微式布局）', async ({ page }) => {
  await page.goto('/');
  // 无列表页：Univer 画布直接挂载
  await expect(page.locator('#univer-container canvas:visible').first()).toBeVisible({
    timeout: 30000,
  });
  // 顶部：logo 品牌名 + 工具栏「菜单」按钮（文档名已移除）
  await expect(page.getByRole('img', { name: '易飞表格' })).toBeVisible();
  await expect(page.getByRole('button', { name: '菜单', exact: true })).toBeVisible();
  // 工具栏：myf-sheet-demo 同款单行工具栏（插入面板 + 常规格式面板）
  await expect(page.getByRole('button', { name: '插入' })).toBeVisible();
  await expect(page.getByRole('button', { name: '常规' })).toBeVisible();
  await page.getByRole('button', { name: '插入' }).click();
  await expect(page.getByText('在上方插入行')).toBeVisible();

  // 菜单含二级「文件」面板与导出项
  await page.getByRole('button', { name: '菜单', exact: true }).click();
  await expect(page.getByText('文件', { exact: true })).toBeVisible();
  await page.getByText('文件', { exact: true }).hover();
  await expect(page.getByText('导出表格(.efexcel)')).toBeVisible();
  await page.keyboard.press('Escape');
});
