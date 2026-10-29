// exceljs Workbook -> Univer snapshot：导入方向的纯映射（不涉及文件读取）
import type { Cell, CellErrorValue, CellFormulaValue, CellHyperlinkValue, CellRichTextValue, CellSharedFormulaValue, Workbook, Worksheet } from 'exceljs';
import type { BorderStyle, CellStyle, MergeRange, SheetRichTextRun, SnapshotCell, SnapshotSheet, WorkbookSnapshot } from '../../types/spreadsheet';
import { newId } from '../../types/spreadsheet';
import { parseThemePalette, resolveColorHex, STANDARD_OFFICE_THEME, type ColorRef } from './colorResolve';

const PT_TO_PX = 1 / 0.75;
const CHAR_TO_PX = 10;

/** exceljs/OOXML 线型 → Univer BorderStyleTypes 数值枚举（1=THIN…13=THICK），两侧词表一致故 1:1 直映 */
const BORDER_STYLE_IDS: Record<string, number> = {
  thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7,
  medium: 8, mediumDashed: 9, mediumDashDot: 10, mediumDashDotDot: 11,
  slantDashDot: 12, thick: 13,
};
/** exceljs 对齐 → Univer HorizontalAlign / VerticalAlign 数值枚举 */
const H_ALIGN_IDS: Record<string, number> = { left: 1, center: 2, right: 3 };
const V_ALIGN_IDS: Record<string, number> = { top: 1, middle: 2, bottom: 3 };
const WRAP_STRATEGY_WRAP = 3; // Univer WrapStrategy.WRAP
/** exceljs font.vertAlign → Univer BaselineOffset（2=下标 3=上标） */
const BASELINE_OFFSET_IDS: Record<string, number> = { subscript: 2, superscript: 3 };

/** Univer 资源名（由各插件在 createWorkbook 时按名反序列化，见 colorResolve 同级的 resources 说明） */
const FILTER_RESOURCE = 'SHEET_FILTER_PLUGIN';
const DEFINED_NAME_RESOURCE = 'SHEET_DEFINED_NAME_PLUGIN';

type BorderSide = { style?: string; color?: ColorRef } | undefined;

/** 读取单边边框：无有效线型时忽略，缺省颜色补黑色 */
function readBorderSide(side: BorderSide, palette: string[]): BorderStyle | undefined {
  const s = side?.style ? BORDER_STYLE_IDS[side.style] : undefined;
  if (!s) return undefined;
  return { s, cl: { rgb: resolveColorHex(side?.color, palette) ?? '#000000' } };
}

/** 从工作簿内嵌的 theme1.xml 解析主题色板；缺失/解析失败回退标准 Office 主题 */
function themePaletteOf(wb: Workbook): string[] {
  const themes = (wb.model as unknown as { themes?: Record<string, string> } | undefined)?.themes;
  const xml = themes ? themes.theme1 ?? Object.values(themes)[0] : undefined;
  return (xml && parseThemePalette(xml)) || STANDARD_OFFICE_THEME;
}

export function workbookToSnapshot(wb: Workbook, name: string): WorkbookSnapshot {
  const palette = themePaletteOf(wb);
  const sheets: Record<string, SnapshotSheet> = {};
  const sheetOrder: string[] = [];
  const filters: Record<string, { ref: MergeRange; filterColumns: never[] }> = {};
  wb.worksheets.forEach((ws, index) => {
    const id = `sheet-${String(index + 1).padStart(2, '0')}`;
    sheets[id] = readSheet(ws, id, palette);
    sheetOrder.push(id);
    const ref = readAutoFilter(ws);
    if (ref) filters[id] = { ref, filterColumns: [] };
  });
  const resources = buildResources(wb, filters);
  return {
    id: newId().replace('doc-', 'workbook-'),
    name,
    sheetOrder,
    sheets,
    ...(resources.length ? { resources } : {}),
  };
}

function readSheet(ws: Worksheet, id: string, palette: string[]): SnapshotSheet {
  const cellData: SnapshotSheet['cellData'] = {};
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    cellData[r - 1] ??= {};
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      const uc = readCell(cell, palette);
      // 富文本单元格仅有 p，也需保留
      if (uc.v !== undefined || uc.f || uc.s || uc.p) cellData[r - 1][c - 1] = uc;
    });
  });

  const mergeData = (sheetMerges(ws)).map((range) => rangeToMerge(range));

  const rowData: SnapshotSheet['rowData'] = {};
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const h = row.height;
    if (row.hidden) rowData[r - 1] = { h: h ? Math.round(h * PT_TO_PX) : 0, hd: 1 };
    else if (h) rowData[r - 1] = { h: Math.round(h * PT_TO_PX) };
  }
  const columnData: SnapshotSheet['columnData'] = {};
  ws.columns?.forEach((col, i) => {
    if (!col) return;
    const w = col.width ? Math.round(col.width * CHAR_TO_PX) : 0;
    if (col.hidden) columnData[i] = { w, hd: 1 };
    else if (w) columnData[i] = { w };
  });

  const sheet: SnapshotSheet = {
    id,
    name: ws.name,
    // veryHidden 与 hidden 同为不可见，导入后均按隐藏处理
    hidden: ws.state === 'hidden' || ws.state === 'veryHidden' ? 1 : undefined,
    rowCount: Math.max(ws.rowCount + 5, 50),
    columnCount: Math.max(ws.columnCount + 3, 20),
    cellData,
    mergeData,
    freeze: readFreeze(ws),
    rowData,
    columnData,
  };

  // 标签页颜色：tabColor 同样可能是主题/索引色，走统一解析
  const tabColor = resolveColorHex(ws.properties?.tabColor as ColorRef | undefined, palette);
  if (tabColor) sheet.tabColor = tabColor;

  // 网格线：视图状态优先（Excel 写在 sheetView），回退 sheet 属性；仅在明确关闭时置 0（缺省即显示）
  const view = ws.views?.[0];
  const showGridLines = view?.showGridLines ?? ws.properties?.showGridLines;
  if (showGridLines === false) sheet.showGridlines = 0;
  if (view?.rightToLeft) sheet.rightToLeft = 1;

  // 默认行高/列宽：无显式尺寸的行列采用；忠实还原源文件声明的默认值（行高 pt→px、列宽字符→px）
  const defaultRowHeight = ws.properties?.defaultRowHeight;
  if (defaultRowHeight) sheet.defaultRowHeight = Math.round(defaultRowHeight * PT_TO_PX);
  const defaultColWidth = ws.properties?.defaultColWidth;
  if (defaultColWidth) sheet.defaultColumnWidth = Math.round(defaultColWidth * CHAR_TO_PX);

  return sheet;
}

/** exceljs 冻结视图 → Univer freeze：xSplit/ySplit 为冻结行列数；部分生产者只写 topLeftCell，此时回退解析 */
function readFreeze(ws: Worksheet): SnapshotSheet['freeze'] {
  const view = ws.views?.[0];
  if (!view || view.state !== 'frozen') return undefined;
  const x = view.xSplit ?? 0;
  const y = view.ySplit ?? 0;
  if (!x && !y) {
    // 回退：仅 topLeftCell 时，其行列号即冻结的列数/行数（如 'B3' → 冻结 1 列 2 行）
    const tl = view.topLeftCell;
    if (!tl) return undefined;
    const { row, col } = decodeRef(tl);
    if (!row && !col) return undefined;
    return { startRow: row, startColumn: col, xSplit: col, ySplit: row };
  }
  return { startRow: y, startColumn: x, xSplit: x, ySplit: y };
}

/** 自动筛选区（ws.autoFilter 为 'A1:D10' 式字符串或 {from,to}）→ Univer IRange；无效引用忽略 */
function readAutoFilter(ws: Worksheet): MergeRange | undefined {
  const af = ws.autoFilter as string | { from?: string; to?: string } | undefined;
  const range = typeof af === 'string' ? af : af?.from ? `${af.from}:${af.to ?? af.from}` : undefined;
  if (!range || !/^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/.test(range)) return undefined;
  return rangeToMerge(range);
}

/** 组装 Univer 插件资源（筛选、定义名称）；均为 JSON 字符串，无内容时不产出对应条目 */
function buildResources(wb: Workbook, filters: Record<string, { ref: MergeRange; filterColumns: never[] }>): { name: string; data: string }[] {
  const resources: { name: string; data: string }[] = [];
  if (Object.keys(filters).length) resources.push({ name: FILTER_RESOURCE, data: JSON.stringify(filters) });
  const definedNames = readDefinedNames(wb);
  if (Object.keys(definedNames).length) resources.push({ name: DEFINED_NAME_RESOURCE, data: JSON.stringify(definedNames) });
  return resources;
}

/** exceljs 定义名称（wb.model.definedNames）→ Univer 定义名称资源（按 id 键控） */
function readDefinedNames(wb: Workbook): Record<string, { id: string; name: string; formulaOrRefString: string }> {
  const list = (wb.model as { definedNames?: { name?: string; ranges?: string[] }[] } | undefined)?.definedNames;
  const out: Record<string, { id: string; name: string; formulaOrRefString: string }> = {};
  if (!Array.isArray(list)) return out;
  for (const dn of list) {
    if (!dn?.name || !dn.ranges?.length) continue;
    out[dn.name] = { id: dn.name, name: dn.name, formulaOrRefString: dn.ranges.join(',') };
  }
  return out;
}

/** exceljs 类型包里 merges 挂在 model 上且可选，做防御式读取 */
function sheetMerges(ws: Worksheet): string[] {
  const model = ws.model as { merges?: string[] } | undefined;
  return model?.merges ?? [];
}

function rangeToMerge(range: string): MergeRange {
  const [tl, br] = range.split(':');
  const a = decodeRef(tl);
  const b = decodeRef(br ?? tl);
  return { startRow: a.row, startColumn: a.col, endRow: b.row, endColumn: b.col };
}

function decodeRef(ref: string): { row: number; col: number } {
  const m = ref.match(/([A-Z]+)(\d+)/);
  if (!m) return { row: 0, col: 0 };
  const col = m[1].split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return { row: Number(m[2]) - 1, col };
}

function readCell(cell: Cell, palette: string[]): SnapshotCell {
  const uc: SnapshotCell = {};
  const val = cell.value;
  if (val != null && typeof val === 'object' && ('formula' in val || 'sharedFormula' in val)) {
    // 普通公式 value 带 formula 文本；共享公式 value 只有 { sharedFormula, result }（formula 为
    // undefined），须读 cell.formula getter——exceljs 已按行列位移把相对引用平移成本单元格的公式
    uc.f = '=' + (cell.formula || (val as CellFormulaValue).formula);
    applyCachedResult(uc, (val as CellFormulaValue | CellSharedFormulaValue).result);
  } else if (val != null && typeof val === 'object' && 'richText' in (val as CellRichTextValue)) {
    // exceljs 类型缺陷：Cell 继承的 Address 把 row/col 误声明为 string，运行时实为 number
    applyRichText(uc, val as CellRichTextValue, Number(cell.row), Number(cell.col), palette);
  } else if (val instanceof Date) {
    uc.v = formatDate(val);
  } else if (val != null && typeof val === 'object' && 'error' in (val as CellErrorValue)) {
    uc.v = String((val as CellErrorValue).error);
  } else if (val != null && typeof val === 'object' && 'hyperlink' in (val as CellHyperlinkValue)) {
    // 超链接对象：显示文本（缺省退化为 URL 本身），不能把对象塞进 v；URL 由 fileIO 侧单独提取落位
    const link = val as CellHyperlinkValue;
    uc.v = typeof link.text === 'string' ? link.text : link.hyperlink;
  } else if (val != null) {
    uc.v = val as string | number | boolean;
  }

  const st = readStyle(cell, palette);
  if (st) uc.s = st;
  return uc;
}

/** 公式缓存结果 → v：仅收原始值；Date/富文本/错误对象等无对应展示形态的不写，避免对象漏进 v 渲染成空白 */
function applyCachedResult(uc: SnapshotCell, result: unknown): void {
  if (result instanceof Date) uc.v = formatDate(result);
  else if (typeof result === 'string' || typeof result === 'number' || typeof result === 'boolean') uc.v = result;
}

function readStyle(cell: Cell, palette: string[]): CellStyle | undefined {
  const st: CellStyle = {};
  const font = cell.style?.font;
  if (font?.bold) st.bl = 1;
  if (font?.italic) st.it = 1;
  if (font?.size != null) st.fs = font.size;
  const fontColor = resolveColorHex(font?.color as ColorRef | undefined, palette);
  if (fontColor) st.cl = { rgb: fontColor };
  if (font?.name) st.ff = font.name;
  if (font?.underline && font.underline !== 'none') st.ul = { s: 1 };
  if (font?.strike) st.st = { s: 1 };
  if (font?.vertAlign && BASELINE_OFFSET_IDS[font.vertAlign]) st.va = BASELINE_OFFSET_IDS[font.vertAlign];
  const fill = cell.style?.fill;
  if (fill && 'pattern' in fill && fill.pattern === 'solid') {
    const bg = resolveColorHex((fill.fgColor ?? fill.bgColor) as ColorRef | undefined, palette);
    if (bg) st.bg = { rgb: bg };
  }
  const align = cell.style?.alignment;
  if (align?.horizontal && H_ALIGN_IDS[align.horizontal] !== undefined) st.ht = H_ALIGN_IDS[align.horizontal];
  if (align?.vertical && V_ALIGN_IDS[align.vertical] !== undefined) st.vt = V_ALIGN_IDS[align.vertical];
  if (align?.wrapText) st.tb = WRAP_STRATEGY_WRAP;
  const rotation = readTextRotation(align?.textRotation);
  if (rotation) st.tr = rotation;
  const bd = cell.style?.border;
  if (bd) {
    const t = readBorderSide(bd.top, palette);
    const b = readBorderSide(bd.bottom, palette);
    const l = readBorderSide(bd.left, palette);
    const r = readBorderSide(bd.right, palette);
    if (t || b || l || r) st.bd = { ...(t && { t }), ...(b && { b }), ...(l && { l }), ...(r && { r }) };
  }
  const numFmt = cell.style?.numFmt;
  if (numFmt) st.n = { pattern: numFmt };
  return Object.keys(st).length ? st : undefined;
}

/**
 * exceljs textRotation → Univer ITextRotation。
 * 'vertical'（OOXML 255，字符竖直堆叠）→ { a: 0, v: 1 }；
 * 数值按 OOXML 约定 0–90 为逆时针（上仰）、91–180 为顺时针（下俯），折算成 Univer 的带符号角度
 * （正=逆时针/上仰，与 Excel 一致；Univer 渲染按 |a| 取大小、按符号定朝向）。
 */
function readTextRotation(rotation: number | 'vertical' | undefined): CellStyle['tr'] {
  if (rotation === 'vertical') return { a: 0, v: 1 };
  if (typeof rotation !== 'number' || rotation === 0) return undefined;
  const a = rotation > 90 ? rotation - 180 : rotation;
  return a === 0 ? undefined : { a };
}

/**
 * Date → 展示串。exceljs 以 UTC 构造 Date（序列号无时区语义），必须用 getUTC* 读取，
 * 否则负时区会把日期读偏一天；含非零时间时补上 HH:MM(:SS)，避免丢掉时间部分。
 */
function formatDate(d: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  const date = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  const h = d.getUTCHours();
  const mi = d.getUTCMinutes();
  const s = d.getUTCSeconds();
  if (!h && !mi && !s) return date;
  return s ? `${date} ${pad(h)}:${pad(mi)}:${pad(s)}` : `${date} ${pad(h)}:${pad(mi)}`;
}

/** exceljs 富文本 → Univer 文档子集：拼接 dataStream（\r\n 为段落分隔，流以 \r\n 结尾）+ 按区间切 textRuns */
function applyRichText(uc: SnapshotCell, val: CellRichTextValue, row: number, col: number, palette: string[]): void {
  const parts = val.richText.map((t) => {
    // 先把 \r\n? 归一化为 \n，再统一展开为 \r\n，避免源文本含 \r\n 时双重展开成 \r\r\n
    const normalized = t.text.replace(/\r\n?/g, '\n');
    return { expanded: normalized.replace(/\n/g, '\r\n'), normalized, ts: t.font ? fontToStyle(t.font, palette) : undefined };
  });
  const dataStream = parts.map((p) => p.expanded).join('') + '\r\n';
  const textRuns: SheetRichTextRun[] = [];
  let offset = 0;
  for (const part of parts) {
    // run 区间按 Univer 惯例以归一化文本计长（\r\n 段落分隔算 1 字符），与 dataStream 的 \r\n 展开无关
    if (part.ts && hasShapeStyle(part.ts)) textRuns.push({ st: offset, ed: offset + part.normalized.length, ts: part.ts });
    offset += part.normalized.length;
  }
  uc.p = {
    // 每单元格唯一 id（含坐标），避免多个富文本单元格共享同一文档 id
    id: `__eflink-rich-text-r${row}-c${col}`,
    body: { dataStream, ...(textRuns.length ? { textRuns } : {}) },
    documentStyle: {},
  };
}

/** exceljs run 字体 → CellStyle（字段名与 CellStyle 有差异，先做显式转换） */
function fontToStyle(font: { bold?: boolean; italic?: boolean; size?: number; name?: string; color?: ColorRef; underline?: boolean | string; strike?: boolean; vertAlign?: string }, palette: string[]): CellStyle {
  const st: CellStyle = {};
  if (font.bold) st.bl = 1;
  if (font.italic) st.it = 1;
  if (font.size != null) st.fs = font.size;
  if (font.name) st.ff = font.name;
  const color = resolveColorHex(font.color, palette);
  if (color) st.cl = { rgb: color };
  if (font.underline && font.underline !== 'none') st.ul = { s: 1 };
  if (font.strike) st.st = { s: 1 };
  if (font.vertAlign && BASELINE_OFFSET_IDS[font.vertAlign]) st.va = BASELINE_OFFSET_IDS[font.vertAlign];
  return st;
}

/** run 是否携带视觉样式（粗/斜/色/下划线/删除线/底色/字号/字体族/上下标）——完全无差异的 run 不产出 */
function hasShapeStyle(ts: CellStyle): boolean {
  return ts.bl !== undefined || ts.it !== undefined || ts.cl !== undefined || ts.ul !== undefined || ts.st !== undefined || ts.bg !== undefined || ts.fs !== undefined || ts.ff !== undefined || ts.va !== undefined;
}
