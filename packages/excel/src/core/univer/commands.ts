// 工具栏 → Univer 操作层：全部走 facade 公开 API 与注册命令
import { Dimension } from '@univerjs/core';
import type { FRange, FWorkbook, FWorksheet } from '@univerjs/preset-sheets-core';
import { getUniverApi } from './controller';
import { buildChartOption, renderChartFrameToDataUrl } from './chartOptions';
import { defaultChartElements, setChartMeta } from './chartSettings';

type Arg<T> = T extends (...args: infer A) => unknown ? A[0] : never;
type Arg2<T> = T extends (...args: infer A) => unknown ? A[1] : never;

export function activeWorkbook(): FWorkbook | null {
  return getUniverApi()?.getActiveWorkbook() ?? null;
}

export function withActiveRange(fn: (range: FRange, wb: FWorkbook) => void): void {
  const wb = activeWorkbook();
  const range = wb?.getActiveRange();
  if (!wb || !range) return;
  fn(range, wb);
}

export function execCommand(id: string, params?: unknown): void {
  void getUniverApi()?.executeCommand(id, params as never);
}

// ---------- 编辑历史 ----------
export const undo = () => activeWorkbook()?.undo();
export const redo = () => activeWorkbook()?.redo();
export const formatPainter = () => execCommand('sheet.command.set-once-format-painter');

// ---------- 字体样式（读取当前样式做 toggle） ----------
function currentStyle(): { bl?: number; it?: number; u?: { s?: boolean }; st?: { s?: boolean } } | null {
  let out: ReturnType<FRange['getCellStyleData']> = null;
  withActiveRange((range) => {
    out = range.getCellStyleData();
  });
  return out as { bl?: number; it?: number; u?: { s?: boolean }; st?: { s?: boolean } } | null;
}

export function toggleBold(): void {
  withActiveRange((range) => {
    const on = currentStyle()?.bl === 1;
    range.setFontWeight((on ? 'normal' : 'bold') as Arg<FRange['setFontWeight']>);
  });
}

export function toggleItalic(): void {
  withActiveRange((range) => {
    const on = currentStyle()?.it === 1;
    range.setFontStyle((on ? 'normal' : 'italic') as Arg<FRange['setFontStyle']>);
  });
}

export function toggleUnderline(): void {
  withActiveRange((range) => {
    const on = currentStyle()?.u?.s === true;
    range.setFontLine((on ? 'none' : 'underline') as Arg<FRange['setFontLine']>);
  });
}

export function toggleStrikeThrough(): void {
  withActiveRange((range) => {
    const on = currentStyle()?.st?.s === true;
    range.setFontLine((on ? 'none' : 'line-through') as Arg<FRange['setFontLine']>);
  });
}

export function setFontFamily(family: string): void {
  withActiveRange((range) => range.setFontFamily(family));
}

export function setFontSize(size: number): void {
  withActiveRange((range) => range.setFontSize(size));
}

// ---------- 颜色 ----------
export function setFontColor(color: string): void {
  withActiveRange((range) => range.setFontColor(color));
}

export function resetFontColor(): void {
  execCommand('sheet.command.reset-range-text-color');
}

export function setBackgroundColor(color: string): void {
  withActiveRange((range) => range.setBackgroundColor(color));
}

export function resetBackgroundColor(): void {
  execCommand('sheet.command.reset-background-color');
}

// ---------- 边框 ----------
type SetBorderFn = FRange['setBorder'];

/** 从运行时取 Univer 枚举（BorderStyleTypes 是数字枚举，需运行时读取） */
function univerEnums(): {
  BorderType: Record<string, Arg<SetBorderFn>>;
  BorderStyleTypes: Record<string, Arg2<SetBorderFn>>;
} | null {
  const api = getUniverApi();
  if (!api) return null;
  return (api as unknown as {
    Enum: {
      BorderType: Record<string, Arg<SetBorderFn>>;
      BorderStyleTypes: Record<string, Arg2<SetBorderFn>>;
    };
  }).Enum;
}

export type BorderTypeKey =
  | 'all' | 'outside' | 'none'
  | 'top' | 'bottom' | 'left' | 'right'
  | 'horizontal' | 'vertical'
  | 'tlbr' | 'bltr';

export type BorderStyleKey = 'thin' | 'medium' | 'dashed' | 'dotted' | 'double';

const BORDER_TYPE_ENUM_KEY: Record<BorderTypeKey, string> = {
  all: 'ALL', outside: 'OUTSIDE', none: 'NONE',
  top: 'TOP', bottom: 'BOTTOM', left: 'LEFT', right: 'RIGHT',
  horizontal: 'HORIZONTAL', vertical: 'VERTICAL',
  tlbr: 'TLBR', bltr: 'BL_TR',
};

const BORDER_STYLE_ENUM_KEY: Record<BorderStyleKey, string> = {
  thin: 'THIN', medium: 'MEDIUM', dashed: 'DASHED', dotted: 'DOTTED', double: 'DOUBLE',
};

/** 给当前选区应用边框（类型 × 线型 × 颜色） */
export function applyBorder(type: BorderTypeKey, style: BorderStyleKey, color?: string): void {
  const en = univerEnums();
  if (!en) return;
  const borderType = en.BorderType[BORDER_TYPE_ENUM_KEY[type]];
  const borderStyle = en.BorderStyleTypes[BORDER_STYLE_ENUM_KEY[style]];
  if (borderType === undefined || borderStyle === undefined) return;
  withActiveRange((range) => {
    if (type === 'none') range.setBorder(borderType, en.BorderStyleTypes.NONE);
    else if (color) range.setBorder(borderType, borderStyle, color);
    else range.setBorder(borderType, borderStyle);
  });
}

// ---------- 对齐 / 换行 / 合并 ----------
export function setHorizontalAlign(value: 'left' | 'center' | 'right'): void {
  withActiveRange((range) => range.setHorizontalAlignment(value as Arg<FRange['setHorizontalAlignment']>));
}

export function setVerticalAlign(value: 'top' | 'middle' | 'bottom'): void {
  withActiveRange((range) => range.setVerticalAlignment(value as Arg<FRange['setVerticalAlignment']>));
}

export function toggleWrap(): void {
  withActiveRange((range) => {
    let wrapped = false;
    try {
      wrapped = range.getWrap();
    } catch {
      wrapped = false;
    }
    range.setWrap(!wrapped);
  });
}

export function mergeCells(): void {
  withActiveRange((range) => range.merge());
}

export function unmergeCells(): void {
  withActiveRange((range) => range.breakApart());
}

/** 合并单元格按钮：选中区已在合并中则取消合并，否则合并 */
export function toggleMergeCells(): void {
  withActiveRange((range) => {
    if (range.isMerged()) range.breakApart();
    else range.merge();
  });
}

// ---------- 插入行列（沿选区边缘插入，插入数量 = 选区行/列数） ----------
export function insertRowAbove(): void {
  withActiveRange((range) => range.insertCells(Dimension.ROWS));
}

export function insertRowBelow(): void {
  withActiveRange((range, wb) => {
    const ws = wb.getActiveSheet();
    ws.getRange(range.getRow() + range.getHeight(), range.getColumn(), range.getHeight(), range.getWidth())
      .insertCells(Dimension.ROWS);
  });
}

export function insertColLeft(): void {
  withActiveRange((range) => range.insertCells(Dimension.COLUMNS));
}

export function insertColRight(): void {
  withActiveRange((range, wb) => {
    const ws = wb.getActiveSheet();
    ws.getRange(range.getRow(), range.getColumn() + range.getWidth(), range.getHeight(), range.getWidth())
      .insertCells(Dimension.COLUMNS);
  });
}

// ---------- 清除 ----------
export function clearContentOnly(): void {
  withActiveRange((range) => range.clearContent());
}

export function clearFormatOnly(): void {
  withActiveRange((range) => range.clearFormat());
}

export function clearAll(): void {
  withActiveRange((range) => range.clear());
}

// ---------- 数字格式 ----------
export function setPercentFormat(): void {
  execCommand('sheet.command.numfmt.set.percent');
}

export function setCurrencyFormat(): void {
  execCommand('sheet.command.numfmt.set.currency');
}

export function addDecimal(): void {
  execCommand('sheet.command.numfmt.add.decimal.command');
}

export function subtractDecimal(): void {
  execCommand('sheet.command.numfmt.subtract.decimal.command');
}

/** 直接按 pattern 设置数字格式（「常规」面板的 15 种格式） */
export function setNumberFormatPattern(pattern: string): void {
  withActiveRange((range) => {
    range.setNumberFormat(pattern);
  });
}

/** 读取选区左上角单元格当前数字格式（无格式返回空串/General） */
export function getActiveNumberFormat(): string {
  let out = '';
  withActiveRange((range) => {
    try {
      out = range.getNumberFormat();
    } catch {
      out = '';
    }
  });
  return out;
}

// ---------- 快速统计（Σ 求和菜单）：在选区正下方写入统计公式 ----------
export type QuickStatKind = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN';

export function applyQuickStat(kind: QuickStatKind): void {
  withActiveRange((range, wb) => {
    const ws = wb.getActiveSheet();
    const target = ws.getRange(range.getRow() + range.getHeight(), range.getColumn());
    target.setFormula(`=${kind}(${range.getA1Notation()})`);
  });
}

// ---------- 筛选（企微「筛选」）：对当前选区开/关筛选，列头出现筛选图标 ----------
export const toggleFilter = (): void => {
  execCommand('sheet.command.smart-toggle-filter');
};

// ---------- 排序（企微「排序」菜单：升序/降序，按活动单元格所在列排序） ----------
// 单个活动单元格 → 自动扩展到连续数据区（企微/Excel 行为）；显式框选区域 → 只对该区域排序
function sortRangeCommand(ascending: boolean): void {
  const range = activeWorkbook()?.getActiveRange();
  const singleCell = !range || (range.getHeight() === 1 && range.getWidth() === 1);
  const ext = singleCell ? '-ext' : '';
  execCommand(`sheet.command.sort-range-${ascending ? 'asc' : 'desc'}${ext}`);
}

export const sortAscending = (): void => sortRangeCommand(true);

export const sortDescending = (): void => sortRangeCommand(false);

/** 自定义排序（Univer 内置面板：多条件 + 「标题不参与排序」） */
export const sortCustom = (): void => {
  execCommand('sheet.command.sort-range-custom');
};

// ---------- 冻结（企微「冻结」菜单） ----------
export function freezeFirstRow(): void {
  getUniverApi()?.getActiveWorkbook()?.getActiveSheet()?.setFrozenRows(1);
}

export function freezeFirstColumn(): void {
  getUniverApi()?.getActiveWorkbook()?.getActiveSheet()?.setFrozenColumns(1);
}

export function freezeToActiveCell(): void {
  withActiveRange((range) => {
    const wb = activeWorkbook();
    const ws = wb?.getActiveSheet();
    if (!ws) return;
    ws.setFrozenRows(range.getRow());
    ws.setFrozenColumns(range.getColumn());
  });
}

export function cancelFreeze(): void {
  const ws = activeWorkbook()?.getActiveSheet();
  if (!ws) return;
  ws.setFrozenRows(0);
  ws.setFrozenColumns(0);
}

// ---------- 插入图片（企微「插入」菜单：单元格图片 / 浮动图片） ----------

/** 浮动图片默认显示尺寸（像素） */
export const FLOATING_IMAGE_DEFAULT_WIDTH = 400;
export const FLOATING_IMAGE_DEFAULT_HEIGHT = 300;

function lastInsertedImage(ws: FWorksheet, beforeCount: number) {
  const images = ws.getImages();
  return images.length > beforeCount ? images[images.length - 1] ?? null : null;
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error ?? new Error('读取图片失败'));
    reader.readAsDataURL(file);
  });
}

/** 单元格图片：嵌入当前活动单元格内，随单元格移动 */
export async function insertCellImage(file: File): Promise<boolean> {
  const range = activeWorkbook()?.getActiveRange();
  if (!range) return false;
  return range.insertCellImageAsync(file);
}

/** 浮动图片：悬浮在表格上层，锚定当前活动单元格，可自由拖动缩放 */
export async function insertFloatingImage(file: File): Promise<boolean> {
  const wb = activeWorkbook();
  const ws = wb?.getActiveSheet();
  if (!wb || !ws) return false;
  const url = await fileToDataUrl(file);
  const range = wb.getActiveRange();
  const col = range?.getColumn() ?? 0;
  const row = range?.getRow() ?? 0;
  const before = ws.getImages().length;
  const inserted = await ws.insertImage(url, col, row);
  if (!inserted) return false;
  const added = lastInsertedImage(ws, before);
  if (!added) return false;
  await added.setSizeAsync(FLOATING_IMAGE_DEFAULT_WIDTH, FLOATING_IMAGE_DEFAULT_HEIGHT);
  return true;
}

// ---------- 插入链接（打开 Univer 内置插入链接面板，作用于当前选区单元格） ----------
export function insertHyperLink(): void {
  execCommand('sheet.operation.insert-hyper-link');
}

// ---------- 插入图表：按当前选区数据离屏渲染（echarts），加边框阴影后以浮动图片插入 ----------

export async function insertChart(typeId: string): Promise<boolean> {
  const wb = activeWorkbook();
  const ws = wb?.getActiveSheet();
  const range = wb?.getActiveRange();
  if (!wb || !ws || !range) return false;
  const matrix = range.getValues() as unknown[][];
  const option = buildChartOption(matrix, typeId);

  const url = await renderChartFrameToDataUrl(option);
  const before = ws.getImages().length;
  const inserted = await ws.insertImage(url, range.getColumn(), range.getRow());
  if (!inserted) return false;
  // 渲染用了 2 倍像素比（保证清晰度），插入后把显示尺寸校正为逻辑尺寸（图表 640x400 + 边距 12x2）
  const added = lastInsertedImage(ws, before);
  if (!added) return false;
  await added.setSizeAsync(664, 424);
  // 登记图表元数据：双击图片时供右侧「图表设置」面板读取与修改
  setChartMeta({
    drawingId: added.getId(),
    typeId,
    matrix,
    elements: defaultChartElements(),
    paletteIndex: 0,
    title: '',
    rangeRef: range.getA1Notation(),
  });
  return true;
}

// ---------- 评论 ----------
export function addComment(): void {
  execCommand('sheet.operation.show-comment-modal');
}

export function toggleCommentPanel(): void {
  execCommand('sheet.operation.toggle-comment-panel');
}

// ---------- 查找替换 ----------
export function openFindReplace(): void {
  execCommand('ui.operation.open-find-dialog');
}

// ---------- 视图 ----------
export function resetZoom(): void {
  execCommand('sheet.command.set-zoom-ratio', { value: 1 });
}

export function resetScroll(): void {
  execCommand('sheet.command.scroll-view-reset');
}

// ---------- 末尾追加行/列（网格角落 ⊕ 按钮）：追加后滚动到新增的行/列 ----------
export function addRowBottom(count = 1): void {
  const ws = activeWorkbook()?.getActiveSheet();
  if (!ws) return;
  ws.insertRowsAfter(ws.getMaxRows() - 1, count);
  ws.scrollToCell(ws.getMaxRows() - 1, 0);
}

export function addColumnRight(count = 1): void {
  const ws = activeWorkbook()?.getActiveSheet();
  if (!ws) return;
  ws.insertColumnsAfter(ws.getMaxColumns() - 1, count);
  ws.scrollToCell(0, ws.getMaxColumns() - 1);
}
