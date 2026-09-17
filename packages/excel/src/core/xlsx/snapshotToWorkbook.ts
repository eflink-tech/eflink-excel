// Univer snapshot -> exceljs Workbook：导出方向的纯映射（不涉及 DOM/下载）
import type { Cell, CellRichTextValue, CellValue, Workbook, Worksheet } from 'exceljs';
import type { CellStyle, SnapshotCell, SnapshotSheet, WorkbookSnapshot } from '../../types/spreadsheet';

const toArgb = (rgb: string): string => 'FF' + rgb.replace('#', '').toUpperCase();
const PX_TO_PT = 0.75; // Univer 行高(px) -> excel 行高(pt)
const PX_TO_CHAR = 10; // Univer 列宽(px) -> excel 字符宽（经验换算）

/** Univer BorderStyleTypes 数值枚举 → exceljs/OOXML 线型（1:1 直映） */
const BORDER_STYLE_NAMES: Record<number, string> = {
  1: 'thin', 2: 'hair', 3: 'dotted', 4: 'dashed', 5: 'dashDot', 6: 'dashDotDot', 7: 'double',
  8: 'medium', 9: 'mediumDashed', 10: 'mediumDashDot', 11: 'mediumDashDotDot',
  12: 'slantDashDot', 13: 'thick',
};
const H_ALIGN_NAMES: Record<number, string> = { 1: 'left', 2: 'center', 3: 'right' };
const V_ALIGN_NAMES: Record<number, string> = { 1: 'top', 2: 'middle', 3: 'bottom' };
const WRAP_STRATEGY_WRAP = 3;

export function snapshotToWorkbook(snapshot: WorkbookSnapshot, wb: Workbook): void {
  const styles = snapshot.styles ?? {};
  const order = snapshot.sheetOrder?.filter((id) => snapshot.sheets[id]) ?? Object.keys(snapshot.sheets);
  for (const sheetId of order) {
    const ws = wb.addWorksheet(snapshot.sheets[sheetId].name);
    writeSheet(snapshot.sheets[sheetId], ws, styles);
  }
}

function writeSheet(sheet: SnapshotSheet, ws: Worksheet, styles: Record<string, CellStyle>): void {
  for (const [r, row] of Object.entries(sheet.cellData)) {
    for (const [c, cell] of Object.entries(row)) {
      const target = ws.getCell(Number(r) + 1, Number(c) + 1);
      target.value = toCellValue(cell);
      applyStyle(target, resolveStyle(cell, styles));
    }
  }
  for (const m of sheet.mergeData ?? []) {
    ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1);
  }
  const freeze = sheet.freeze;
  if (freeze && (freeze.startRow || freeze.startColumn)) {
    // Univer freeze（startRow/startColumn=冻结行列数）→ exceljs 冻结视图（ySplit/xSplit 同语义）
    ws.views = [{
      state: 'frozen',
      xSplit: freeze.startColumn,
      ySplit: freeze.startRow,
      topLeftCell: encodeRef(freeze.startRow, freeze.startColumn),
    }] as never;
  }
  for (const [r, rd] of Object.entries(sheet.rowData ?? {})) {
    if (rd.hd) ws.getRow(Number(r) + 1).hidden = true;
    else if (rd.h) ws.getRow(Number(r) + 1).height = Math.max(6, rd.h * PX_TO_PT);
  }
  for (const [c, cd] of Object.entries(sheet.columnData ?? {})) {
    if (cd.hd) ws.getColumn(Number(c) + 1).hidden = true;
    else if (cd.w) ws.getColumn(Number(c) + 1).width = Math.max(5, cd.w / PX_TO_CHAR);
  }
}

/** Univer 运行时快照的 cell.s 是样式 id（styles 表键），导入方向快照则是内联样式对象 */
function resolveStyle(cell: SnapshotCell, styles: Record<string, CellStyle>): CellStyle | undefined {
  if (typeof cell.s === 'string') return styles[cell.s];
  return cell.s;
}

function toCellValue(cell: SnapshotCell): CellValue {
  if (cell.p) return richTextToCellValue(cell.p) as never;
  if (cell.f) return { formula: cell.f.replace(/^=/, ''), result: (cell.v ?? 0) as number };
  return cell.v ?? null;
}

/** Univer 文档子集 → exceljs 富文本：流去段落符还原文本，runs 覆盖区间外的字符输出无样式片段 */
function richTextToCellValue(p: { body: { dataStream: string; textRuns?: { st: number; ed: number; ts?: CellStyle }[] } }): CellRichTextValue {
  const stream = p.body.dataStream.replace(/\r\n/g, '\n').replace(/\n$/, '');
  const runs = [...(p.body.textRuns ?? [])].sort((a, b) => a.st - b.st);
  const richText: { text: string; font?: Record<string, unknown> }[] = [];
  let cursor = 0;
  const pushRun = (text: string, ts?: CellStyle) => {
    if (!text) return;
    richText.push({
      text,
      ...(ts && {
        font: {
          ...(ts.bl != null && { bold: ts.bl === 1 }),
          ...(ts.it != null && { italic: ts.it === 1 }),
          ...(ts.fs != null && { size: ts.fs }),
          ...(ts.ff && { name: ts.ff }),
          ...(ts.cl && { color: { argb: toArgb(ts.cl.rgb) } }),
          ...(ts.ul != null && { underline: ts.ul.s === 1 }),
          ...(ts.st != null && { strike: ts.st.s === 1 }),
        },
      }),
    });
  };
  for (const run of runs) {
    if (run.st > cursor) pushRun(stream.slice(cursor, run.st)); // 无样式间隙
    pushRun(stream.slice(Math.max(run.st, cursor), run.ed), run.ts);
    cursor = Math.max(cursor, run.ed);
  }
  pushRun(stream.slice(cursor));
  return { richText: richText as CellRichTextValue['richText'] };
}

function applyStyle(target: Cell, st: CellStyle | undefined): void {
  if (!st) return;
  if (st.n?.pattern) target.numFmt = st.n.pattern;
  target.font = {
    bold: st.bl === 1,
    italic: st.it === 1,
    size: st.fs ?? 11,
    ...(st.ff ? { name: st.ff } : {}),
    ...(st.cl ? { color: { argb: toArgb(st.cl.rgb) } } : {}),
    ...(st.ul?.s ? { underline: true } : {}),
    ...(st.st?.s ? { strike: true } : {}),
  };
  if (st.bg) {
    target.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: toArgb(st.bg.rgb) } };
  }
  const alignment: Record<string, unknown> = {};
  if (st.ht && H_ALIGN_NAMES[st.ht]) alignment.horizontal = H_ALIGN_NAMES[st.ht];
  if (st.vt && V_ALIGN_NAMES[st.vt]) alignment.vertical = V_ALIGN_NAMES[st.vt];
  if (st.tb === WRAP_STRATEGY_WRAP) alignment.wrapText = true;
  if (Object.keys(alignment).length) target.alignment = alignment as never;
  const bd = st.bd;
  if (bd) {
    const border: Record<string, unknown> = {};
    const sides = [
      ['top', bd.t], ['bottom', bd.b], ['left', bd.l], ['right', bd.r],
    ] as const;
    for (const [key, side] of sides) {
      if (!side) continue;
      const style = BORDER_STYLE_NAMES[side.s];
      if (!style) continue;
      border[key] = { style, color: { argb: toArgb(side.cl.rgb) } };
    }
    if (Object.keys(border).length) target.border = border as never;
  }
}

/** 0 基行列 → 'A1' 式引用（冻结视图 topLeftCell 用） */
function encodeRef(row: number, col: number): string {
  let s = '';
  let c = col;
  while (c >= 0) {
    s = String.fromCharCode((c % 26) + 65) + s;
    c = Math.floor(c / 26) - 1;
  }
  return `${s}${row + 1}`;
}
