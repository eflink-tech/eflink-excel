// Univer snapshot -> exceljs Workbook：导出方向的纯映射（不涉及 DOM/下载）
import type { Cell, CellValue, Workbook, Worksheet } from 'exceljs';
import type { CellStyle, SnapshotCell, SnapshotSheet, WorkbookSnapshot } from '../../types/spreadsheet';

const toArgb = (rgb: string): string => 'FF' + rgb.replace('#', '').toUpperCase();
const PX_TO_PT = 0.75; // Univer 行高(px) -> excel 行高(pt)
const PX_TO_CHAR = 10; // Univer 列宽(px) -> excel 字符宽（经验换算）

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
  for (const [r, rd] of Object.entries(sheet.rowData ?? {})) {
    if (rd.h) ws.getRow(Number(r) + 1).height = Math.max(6, rd.h * PX_TO_PT);
  }
  for (const [c, cd] of Object.entries(sheet.columnData ?? {})) {
    if (cd.w) ws.getColumn(Number(c) + 1).width = Math.max(5, cd.w / PX_TO_CHAR);
  }
}

/** Univer 运行时快照的 cell.s 是样式 id（styles 表键），导入方向快照则是内联样式对象 */
function resolveStyle(cell: SnapshotCell, styles: Record<string, CellStyle>): CellStyle | undefined {
  if (typeof cell.s === 'string') return styles[cell.s];
  return cell.s;
}

function toCellValue(cell: SnapshotCell): CellValue {
  if (cell.f) return { formula: cell.f.replace(/^=/, ''), result: (cell.v ?? 0) as number };
  return cell.v ?? null;
}

function applyStyle(target: Cell, st: CellStyle | undefined): void {
  if (!st) return;
  if (st.n?.pattern) target.numFmt = st.n.pattern;
  target.font = {
    bold: st.bl === 1,
    italic: st.it === 1,
    size: st.fs ?? 11,
    ...(st.cl ? { color: { argb: toArgb(st.cl.rgb) } } : {}),
  };
  if (st.bg) {
    target.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: toArgb(st.bg.rgb) } };
  }
}
