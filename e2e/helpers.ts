import { expect, type Page } from '@playwright/test';

/** 行头列宽 / 列头行高（100% 缩放时的 Univer 默认取样，与 GridEdgeControls 一致） */
export const ROW_HEADER_WIDTH = 46;
export const COL_HEADER_HEIGHT = 24;
/** 默认行高（px），由点击 A2 的坐标实测得出 */
export const ROW_HEIGHT = 24;

export const nameBox = (page: Page) => page.getByTitle('当前单元格');
export const contentInput = (page: Page) => page.getByPlaceholder('输入内容或公式，回车确认');

/** Univer 隐藏编辑器（接收键盘输入的 contenteditable） */
function editorFocused(page: Page): Promise<boolean> {
  return page.evaluate(() => document.activeElement?.getAttribute('data-u-comp') === 'editor');
}

async function canvasA1Point(page: Page): Promise<{ x: number; y: number }> {
  const box = await page.locator('#univer-container').boundingBox();
  if (!box) throw new Error('未找到编辑器容器');
  return { x: box.x + ROW_HEADER_WIDTH + 30, y: box.y + COL_HEADER_HEIGHT + 12 };
}

/**
 * 点击 A1 并确认键盘方向键可用。
 * 无头 Chromium 下点击"非当前选中格"后 Univer 会丢失键盘路由（真实浏览器无此问题），
 * 因此点击后先探测 ArrowDown 是否真正移动选区，失效则对同一格补一次点击（间隔 >600ms 避免双击）。
 */
export async function clickCellA1(page: Page): Promise<void> {
  const { x, y } = await canvasA1Point(page);
  for (let i = 0; i < 5; i++) {
    await page.mouse.click(x, y);
    await expect
      .poll(() => editorFocused(page), { timeout: 5000, intervals: [100] })
      .toBe(true);
    await expect(nameBox(page)).toHaveValue('A1');
    await page.keyboard.press('ArrowDown');
    await page.waitForTimeout(500);
    if ((await nameBox(page).inputValue()) === 'A2') {
      await page.keyboard.press('ArrowUp');
      await page.waitForTimeout(300);
      await expect(nameBox(page)).toHaveValue('A1');
      return;
    }
    await page.waitForTimeout(600);
  }
  throw new Error('画布点击后键盘导航未恢复');
}

/** 打开应用并等待 Univer 画布、公式栏与键盘导航全部就绪（默认选区 A1） */
export async function waitEditor(page: Page): Promise<void> {
  await page.goto('/');
  await expect(page.locator('#univer-container canvas:visible').first()).toBeVisible({
    timeout: 30000,
  });
  await expect(nameBox(page)).toHaveValue('A1', { timeout: 15000 });
  await clickCellA1(page);
}

/** ☰ 主菜单 → 一级分组 → 二级动作（hover 展开二级面板后点击） */
export async function menuClick(page: Page, group: string, leaf?: string): Promise<void> {
  await page.getByRole('button', { name: '菜单' }).click();
  const groupBtn = page.getByRole('button', { name: group, exact: true });
  if (leaf) {
    await groupBtn.hover();
    await page.getByRole('button', { name: leaf, exact: true }).click();
  } else {
    await groupBtn.click();
  }
}

/** 在公式栏输入并回车提交（写入当前选区） */
export async function typeIntoFormulaBar(page: Page, text: string): Promise<void> {
  await contentInput(page).fill(text);
  await contentInput(page).press('Enter');
}

/** 断言当前选区（名称框）与内容框回显 */
export async function expectSelection(page: Page, notation: string, value?: string): Promise<void> {
  await expect(nameBox(page)).toHaveValue(notation);
  if (value !== undefined) {
    await expect(contentInput(page)).toHaveValue(value);
  }
}
