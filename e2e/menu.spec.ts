import { expect, test } from '@playwright/test';
import {
  clickCellA1,
  expectSelection,
  menuClick,
  nameBox,
  typeIntoFormulaBar,
  waitEditor,
} from './helpers';

test.describe('主菜单功能', () => {
  test('快捷键弹窗：打开、Escape 关闭、按钮关闭', async ({ page }) => {
    await waitEditor(page);

    await menuClick(page, '帮助', '快捷键');
    await expect(page.getByRole('heading', { name: '快捷键' })).toBeVisible();
    await expect(page.getByText('保存到本地数据库')).toBeVisible();

    await page.keyboard.press('Escape');
    await expect(page.getByRole('heading', { name: '快捷键' })).toBeHidden();

    await menuClick(page, '帮助', '快捷键');
    await expect(page.getByRole('heading', { name: '快捷键' })).toBeVisible();
    await page.getByRole('button', { name: '关闭' }).click();
    await expect(page.getByRole('heading', { name: '快捷键' })).toBeHidden();
  });

  test('新建表格：二次确认后切换到新文档（localStorage 记录更新、内容为空）', async ({ page }) => {
    await waitEditor(page);

    await typeIntoFormulaBar(page, '旧文档内容');
    const lastDocBefore = await page.evaluate(() => localStorage.getItem('doculuma:excel:last-doc'));

    // 第一次：弹窗出现后取消，不新建
    await menuClick(page, '文件', '新建表格');
    const dialog = page.getByRole('dialog', { name: '新建文档' });
    await expect(dialog).toBeVisible();
    await expect(page.getByText('新建文档将清空当前内容，请确保已经下载备份文档！')).toBeVisible();
    await dialog.getByRole('button', { name: '取消' }).click();
    await expect(dialog).toBeHidden();
    expect(await page.evaluate(() => localStorage.getItem('doculuma:excel:last-doc'))).toBe(lastDocBefore);

    // 第二次：确认新建
    await menuClick(page, '文件', '新建表格');
    const dialog2 = page.getByRole('dialog', { name: '新建文档' });
    await expect(dialog2).toBeVisible();
    await dialog2.getByRole('button', { name: '确定' }).click();
    // 等待新文档落库并写入 lastDoc（避免在旧文档页面上过早断言）
    await expect
      .poll(() => page.evaluate(() => localStorage.getItem('doculuma:excel:last-doc')), { timeout: 10000 })
      .not.toBe(lastDocBefore);
    await expect(nameBox(page)).toHaveValue('A1', { timeout: 15000 });

    // 新文档 A1 为空
    await clickCellA1(page);
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowUp');
    await expectSelection(page, 'A1', '');
  });

  test('冻结首行 / 冻结首列 / 取消冻结动作可执行且画布正常', async ({ page }) => {
    await waitEditor(page);

    await menuClick(page, '视图', '冻结首行');
    await expect(page.locator('#univer-container canvas:visible').first()).toBeVisible();

    await menuClick(page, '视图', '冻结首列');
    await expect(page.locator('#univer-container canvas:visible').first()).toBeVisible();

    await menuClick(page, '视图', '取消冻结');
    await expect(page.locator('#univer-container canvas:visible').first()).toBeVisible();
    // 冻结操作不影响单元格选择与内容编辑
    await clickCellA1(page);
    await typeIntoFormulaBar(page, '冻结后仍可编辑');
    await clickCellA1(page);
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowUp');
    await expectSelection(page, 'A1', '冻结后仍可编辑');
  });

  test('Σ 快速统计：对 A1:A2 求和写入 A3', async ({ page }) => {
    await waitEditor(page);

    await typeIntoFormulaBar(page, '1');
    await clickCellA1(page);
    await page.keyboard.press('ArrowDown');
    await typeIntoFormulaBar(page, '2');

    // 选中 A1:A2 → 工具栏统计 → 求和
    await clickCellA1(page);
    await page.keyboard.press('Shift+ArrowDown');
    await expect(nameBox(page)).toHaveValue('A1:A2');
    await page.getByRole('button', { name: '统计' }).click();
    await page.getByText('求和').click();

    // 统计公式写入选区正下方 A3
    await clickCellA1(page);
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await expectSelection(page, 'A3', '=SUM(A1:A2)');
  });

  test('工具栏数字格式：百分比应用于 A1', async ({ page }) => {
    await waitEditor(page);

    await typeIntoFormulaBar(page, '0.95');
    await clickCellA1(page);
    await page.getByRole('button', { name: '数字格式' }).click();
    await page.getByText('百分比').click();
    await page.keyboard.press('Escape');

    // 公式栏回显格式化后的显示值，画布保持可用
    await expect(page.locator('#univer-container canvas:visible').first()).toBeVisible();
    await clickCellA1(page);
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowUp');
    await expectSelection(page, 'A1', '95.00%');
  });
});

test.describe('网格边缘追加行列', () => {
  // 高度需容纳 100 行全部进入画布（内容末缘不超出可视区，⊕ 才会显示）
  test.use({ viewport: { width: 1900, height: 3200 } });

  test('底部 ⊕ 添加一行后可导航到新增行', async ({ page }) => {
    await waitEditor(page);

    // 默认 100 行，大视口下内容末缘进入可视区，⊕ 可见
    const addRow = page.getByRole('button', { name: '在底部添加行' });
    await expect(addRow).toBeVisible({ timeout: 15000 });
    await addRow.click();

    await clickCellA1(page);
    for (let i = 0; i < 100; i++) {
      await page.keyboard.press('ArrowDown');
    }
    await expect(nameBox(page)).toHaveValue('A101', { timeout: 10000 });
  });

  test('右侧 ⊕ 添加一列后可导航到新增列', async ({ page }) => {
    await waitEditor(page);

    const addCol = page.getByRole('button', { name: '在右侧添加列' });
    await expect(addCol).toBeVisible({ timeout: 15000 });
    await addCol.click();

    // 覆盖带只能落在内容末缘之后：若从表格中间开始，整表看起来只剩一列
    const rightBandLefts = await page.evaluate(() => {
      const canvas = document.querySelector('canvas[id^="univer-sheet-main-canvas"]')!.getBoundingClientRect();
      return [...document.querySelectorAll('div.absolute')]
        .filter((el) => getComputedStyle(el).backgroundColor === 'rgb(246, 246, 247)')
        .map((el) => el.getBoundingClientRect())
        // 右侧覆盖带从画布顶部开始，底部覆盖带从内容末行开始
        .filter((r) => r.width > 0 && r.height > 0 && Math.abs(r.top - canvas.top) < 10)
        .map((r) => r.left - canvas.right);
    });
    expect(rightBandLefts.every((d) => d > -60)).toBe(true);

    await clickCellA1(page);
    for (let i = 0; i < 20; i++) {
      await page.keyboard.press('ArrowRight');
    }
    // 默认 20 列（A~T），新增一列后第 21 列为 U
    await expect(nameBox(page)).toHaveValue('U1', { timeout: 10000 });
  });

  test('右键菜单输入列数后回车即可插入', async ({ page }) => {
    await waitEditor(page);
    const box = await page.locator('#univer-container').boundingBox();
    if (!box) throw new Error('未找到编辑器容器');
    // 右键列头 K（第 11 列）
    await page.mouse.click(box.x + 46 + 10 * 88 + 44, box.y + 12, { button: 'right' });

    const row = page.getByRole('button').filter({ hasText: '在右侧插入' }).filter({ visible: true }).first();
    const countInput = row.locator('input');
    await expect(countInput).toBeVisible({ timeout: 10000 });
    await countInput.fill('12');
    await countInput.press('Enter');

    // 默认 20 列（A~T），在 K 右侧插入 12 列后共 32 列，末列为 AF
    await clickCellA1(page);
    for (let i = 0; i < 31; i++) {
      await page.keyboard.press('ArrowRight');
    }
    await expect(nameBox(page)).toHaveValue('AF1', { timeout: 10000 });
  });

  test('AI 助手面板打开后不被内容外侧覆盖带遮住', async ({ page }) => {
    await waitEditor(page);
    await page.getByTestId('ai-toggle').click();
    const panel = page.getByTestId('ai-panel');
    await expect(panel).toBeVisible();
    // 等宽度过渡与几何重算稳定
    await page.waitForTimeout(600);

    // 覆盖带（#f6f6f7）在 AI 面板开合后必须钳在网格列内，不能压到面板上
    const spill = await page.evaluate(() => {
      const panelRect = document.querySelector('[data-testid="ai-panel"]')!.getBoundingClientRect();
      return [...document.querySelectorAll('div.absolute')]
        .filter((el) => getComputedStyle(el).backgroundColor === 'rgb(246, 246, 247)')
        .some((el) => el.getBoundingClientRect().right > panelRect.left + 1);
    });
    expect(spill).toBe(false);
    // 面板自身内容可见（头部标题不被盖住）
    await expect(panel.getByText('AI 助手').first()).toBeVisible();
  });
});
