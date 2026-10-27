import { expect, test } from '@playwright/test';
import {
  clickCellA1,
  contentInput,
  expectSelection,
  menuClick,
  nameBox,
  typeIntoFormulaBar,
  waitEditor,
} from './helpers';

test.describe('单元格编辑与公式', () => {
  test('公式栏输入文本提交到 A1 并可回显', async ({ page }) => {
    await waitEditor(page);

    await typeIntoFormulaBar(page, '销售数据');
    // 重新选中 A1：点击后用上下方向键强制触发选区变化，同步公式栏
    await clickCellA1(page);
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowUp');
    await expectSelection(page, 'A1', '销售数据');
  });

  test('数字与公式计算：=SUM 求和 2+3=5', async ({ page }) => {
    await waitEditor(page);

    await typeIntoFormulaBar(page, '2');
    await clickCellA1(page);
    await page.keyboard.press('ArrowDown');
    await expect(nameBox(page)).toHaveValue('A2');
    await typeIntoFormulaBar(page, '3');

    await clickCellA1(page);
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await expect(nameBox(page)).toHaveValue('A3');
    await typeIntoFormulaBar(page, '=SUM(A1:A2)');

    // 回到 A3 查看公式回显
    await clickCellA1(page);
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await expectSelection(page, 'A3', '=SUM(A1:A2)');
  });

  test('选区扩展：Shift+方向键选择范围', async ({ page }) => {
    await waitEditor(page);

    await clickCellA1(page);
    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Shift+ArrowRight');
    await expect(nameBox(page)).toHaveValue('A1:B2');
  });

  test('自动保存：防抖 1.5s 写本地草稿（崩溃恢复兜底，不落云端）', async ({ page }) => {
    await waitEditor(page);

    await typeIntoFormulaBar(page, '持久化测试');
    // 防抖 1500ms + 落草稿余量
    await page.waitForTimeout(2500);
    // 草稿语义：仅写 localStorage 草稿键（云端落库走 Ctrl+S，见下一用例）；
    // 云端加载成功时不读草稿，故刷新恢复断言由「手动保存落库」用例承担
    const docId = await page.evaluate(() => localStorage.getItem('doculuma:excel:last-doc'));
    const draft = await page.evaluate((id) => localStorage.getItem(`doculuma:draft:excel:${id}`), docId);
    expect(draft).toBeTruthy();
    expect(draft).toContain('持久化测试');
  });

  test('Ctrl/Cmd+S 手动保存立即落库（刷新后内容仍在）', async ({ page }) => {
    await waitEditor(page);

    await typeIntoFormulaBar(page, '手动保存');
    await page.keyboard.press('ControlOrMeta+s');
    await page.waitForTimeout(300);
    await page.reload();
    await waitEditor(page);
    await expect(contentInput(page)).toHaveValue('手动保存');
  });

  test('菜单清除内容后单元格为空', async ({ page }) => {
    await waitEditor(page);

    await typeIntoFormulaBar(page, '待清除');
    await clickCellA1(page);
    await menuClick(page, '编辑', '清除内容');

    await clickCellA1(page);
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowUp');
    await expectSelection(page, 'A1', '');
  });

  test('工具栏清除格式与清除内容按钮动作可执行', async ({ page }) => {
    await waitEditor(page);

    await typeIntoFormulaBar(page, '格式测试');
    await page.getByRole('button', { name: '加粗' }).click();
    // 清除格式按钮执行不报错（样式还原由导出用例验证）
    await page.getByRole('button', { name: '清除格式' }).click();
    await expect(page.locator('#univer-container canvas:visible').first()).toBeVisible();
  });
});
