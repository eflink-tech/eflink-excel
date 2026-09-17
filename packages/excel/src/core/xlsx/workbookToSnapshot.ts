// exceljs Workbook -> Univer snapshot：导入方向的纯映射（不涉及文件读取）
import type { Cell, CellErrorValue, CellFormulaValue, CellRichTextValue, Workbook, Worksheet } from 'exceljs';
import type { BorderStyle, CellStyle, MergeRange, SheetRichTextRun, SnapshotCell, SnapshotSheet, WorkbookSnapshot } from '../../types/spreadsheet';
import { newId } from '../../types/spreadsheet';

/** ARGB → #RRGGBB（统一小写，ExcelJS 保留输入原样大小写） */
const fromArgb = (argb: string): string => '#' + argb.slice(2).toLowerCase();
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

type BorderSide = { style?: string; color?: { argb?: string } } | undefined;

/** 读取单边边框：无有效线型时忽略，缺省颜色补黑色 */
function readBorderSide(side: BorderSide): BorderStyle | undefined {
  const s = side?.style ? BORDER_STYLE_IDS[side.style] : undefined;
  if (!s) return undefined;
  return { s, cl: { rgb: side?.color?.argb ? fromArgb(side.color.argb) : '#000000' } };
}

export function workbookToSnapshot(wb: Workbook, name: string): WorkbookSnapshot {
  const sheets: Record<string, SnapshotSheet> = {};
  const sheetOrder: string[] = [];
  wb.worksheets.forEach((ws, index) => {
    const id = `sheet-${String(index + 1).padStart(2, '0')}`;
    sheets[id] = readSheet(ws, id);
    sheetOrder.push(id);
  });
  return { id: newId().replace('doc-', 'workbook-'), name, sheetOrder, sheets };
}

function readSheet(ws: Worksheet, id: string): SnapshotSheet {
  const cellData: SnapshotSheet['cellData'] = {};
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    cellData[r - 1] ??= {};
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      const uc = readCell(cell);
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

  return {
    id,
    name: ws.name,
    rowCount: Math.max(ws.rowCount + 5, 50),
    columnCount: Math.max(ws.columnCount + 3, 20),
    cellData,
    mergeData,
    freeze: readFreeze(ws),
    rowData,
    columnData,
  };
}

/** exceljs 冻结视图（xSplit/ySplit=冻结行列数）→ Univer freeze（startRow/startColumn=滚动区起点） */
function readFreeze(ws: Worksheet): SnapshotSheet['freeze'] {
  const view = ws.views?.[0] as { state?: string; xSplit?: number; ySplit?: number } | undefined;
  if (!view || view.state !== 'frozen') return undefined;
  const x = view.xSplit ?? 0;
  const y = view.ySplit ?? 0;
  if (!x && !y) return undefined;
  return { startRow: y, startColumn: x, xAxisSplit: x, yAxisSplit: y };
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

function readCell(cell: Cell): SnapshotCell {
  const uc: SnapshotCell = {};
  const val = cell.value;
  if (val != null && typeof val === 'object' && 'formula' in (val as CellFormulaValue)) {
    uc.f = '=' + (val as CellFormulaValue).formula;
    const result = (val as CellFormulaValue).result;
    if (result != null) uc.v = result as string | number | boolean;
  } else if (val != null && typeof val === 'object' && 'richText' in (val as CellRichTextValue)) {
    applyRichText(uc, val as CellRichTextValue);
  } else if (val instanceof Date) {
    uc.v = formatDate(val);
  } else if (val != null && typeof val === 'object' && 'error' in (val as CellErrorValue)) {
    uc.v = String((val as CellErrorValue).error);
  } else if (val != null) {
    uc.v = val as string | number | boolean;
  }

  const st = readStyle(cell);
  if (st) uc.s = st;
  return uc;
}

function readStyle(cell: Cell): CellStyle | undefined {
  const st: CellStyle = {};
  const font = cell.style?.font;
  if (font?.bold) st.bl = 1;
  if (font?.italic) st.it = 1;
  if (font?.size != null) st.fs = font.size;
  if (font?.color?.argb) st.cl = { rgb: fromArgb(font.color.argb) };
  if (font?.name) st.ff = font.name;
  if (font?.underline && font.underline !== 'none') st.ul = { s: 1 };
  if (font?.strike) st.st = { s: 1 };
  const fill = cell.style?.fill;
  if (fill && 'pattern' in fill && fill.pattern === 'solid') {
    const fg = fill.fgColor?.argb;
    if (fg) st.bg = { rgb: fromArgb(fg) };
  }
  const align = cell.style?.alignment;
  if (align?.horizontal && H_ALIGN_IDS[align.horizontal] !== undefined) st.ht = H_ALIGN_IDS[align.horizontal];
  if (align?.vertical && V_ALIGN_IDS[align.vertical] !== undefined) st.vt = V_ALIGN_IDS[align.vertical];
  if (align?.wrapText) st.tb = WRAP_STRATEGY_WRAP;
  const bd = cell.style?.border;
  if (bd) {
    const t = readBorderSide(bd.top);
    const b = readBorderSide(bd.bottom);
    const l = readBorderSide(bd.left);
    const r = readBorderSide(bd.right);
    if (t || b || l || r) st.bd = { ...(t && { t }), ...(b && { b }), ...(l && { l }), ...(r && { r }) };
  }
  const numFmt = cell.style?.numFmt;
  if (numFmt) st.n = { pattern: numFmt };
  return Object.keys(st).length ? st : undefined;
}

function formatDate(d: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

const RICH_TEXT_DOC_ID = '__eflink-rich-text';

/** exceljs 富文本 → Univer 文档子集：拼接 dataStream（\r\n 为段落分隔，流以 \r\n 结尾）+ 按区间切 textRuns */
function applyRichText(uc: SnapshotCell, val: CellRichTextValue): void {
  const parts = val.richText.map((t) => ({ text: t.text.replace(/\n/g, '\r\n'), ts: t.font ? fontToStyle(t.font) : undefined }));
  const dataStream = parts.map((p) => p.text).join('') + '\r\n';
  const textRuns: SheetRichTextRun[] = [];
  let offset = 0;
  for (const part of parts) {
    // 纯字体度量差异（仅 fs/ff）不产出 run，避免普通文本碎片化；仅形状样式才切 run
    if (part.ts && hasShapeStyle(part.ts)) textRuns.push({ st: offset, ed: offset + part.text.length, ts: part.ts });
    offset += part.text.length;
  }
  uc.p = {
    id: RICH_TEXT_DOC_ID,
    body: { dataStream, ...(textRuns.length && { textRuns }) },
    documentStyle: {},
  };
}

/** exceljs run 字体 → CellStyle（字段名与 CellStyle 有差异，先做显式转换） */
function fontToStyle(font: { bold?: boolean; italic?: boolean; size?: number; name?: string; color?: { argb?: string }; underline?: boolean | string; strike?: boolean }): CellStyle {
  const st: CellStyle = {};
  if (font.bold) st.bl = 1;
  if (font.italic) st.it = 1;
  if (font.size != null) st.fs = font.size;
  if (font.name) st.ff = font.name;
  if (font.color?.argb) st.cl = { rgb: fromArgb(font.color.argb) };
  if (font.underline && font.underline !== 'none') st.ul = { s: 1 };
  if (font.strike) st.st = { s: 1 };
  return st;
}

/** run 是否携带形状样式（粗/斜/色/下划线/删除线/底色）——仅 fs/ff 度量差异不算 */
function hasShapeStyle(ts: CellStyle): boolean {
  return ts.bl !== undefined || ts.it !== undefined || ts.cl !== undefined || ts.ul !== undefined || ts.st !== undefined || ts.bg !== undefined;
}
