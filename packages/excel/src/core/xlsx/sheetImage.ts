// 当前 sheet 整表导出为 PNG：按已用区域重建一张 HTML 表格再截图，
// 而不是截编辑器可视区域（长表/宽表会丢内容）。显示值由 Univer facade 提供，
// 保证公式结果、数字格式与界面一致。
import html2canvas from 'html2canvas';
import { downloadBlob } from './fileIO';
import type { CellStyle, SnapshotCell, SnapshotSheet } from '../../types/spreadsheet';

export interface SheetImageData {
  sheet: SnapshotSheet;
  styles?: Record<string, CellStyle>;
  /** Univer getDisplayValues：已应用公式结果与数字格式 */
  displayValues: string[][];
  /** Univer getValues：仅用于推断默认对齐（数字右对齐等） */
  rawValues: unknown[][];
  rowHeights: number[];
  columnWidths: number[];
}

/** Univer BorderStyleTypes → CSS 线型（近似值，双层/粗线保留视觉差异） */
const BORDER_CSS: Record<number, string> = {
  1: '1px solid',
  2: '1px solid',
  3: '1px dotted',
  4: '1px dashed',
  5: '1px dashed',
  6: '1px dashed',
  7: '3px double',
  8: '2px solid',
  9: '2px dashed',
  10: '2px dashed',
  11: '2px dashed',
  12: '1px dashed',
  13: '3px solid',
};

const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COLUMN_WIDTH = 88;

function styleOf(data: SheetImageData, cell?: SnapshotCell): CellStyle | undefined {
  if (!cell?.s) return undefined;
  return typeof cell.s === 'string' ? data.styles?.[cell.s] : cell.s;
}

/** 富文本兜底：displayValue 为空时读 textRun 的纯文本 */
function richTextOf(cell?: SnapshotCell): string {
  const stream = cell?.p?.body?.dataStream ?? '';
  return stream.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n+$/, '');
}

function defaultAlign(cell: SnapshotCell | undefined, raw: unknown): string {
  if (cell?.p) return 'left';
  if (typeof raw === 'number') return 'right';
  if (typeof raw === 'boolean') return 'center';
  return 'left';
}

function resolveFontColor(style: CellStyle | undefined): string | undefined {
  const rgb = style?.cl?.rgb;
  return rgb && /^(#|rgb)/.test(rgb) ? rgb : undefined;
}

function resolveBackground(style: CellStyle | undefined): string | undefined {
  const rgb = style?.bg?.rgb;
  return rgb && /^(#|rgb)/.test(rgb) ? rgb : undefined;
}

function borderOf(style: CellStyle | undefined, side: 't' | 'r' | 'b' | 'l'): string | undefined {
  const border = style?.bd?.[side];
  if (!border) return undefined;
  const css = BORDER_CSS[border.s];
  if (!css) return undefined;
  const color = border.cl?.rgb && /^(#|rgb)/.test(border.cl.rgb) ? border.cl.rgb : '#333';
  return `${css} ${color}`;
}

/** 按当前 sheet 已用区域生成整表 DOM（纯函数，便于单测） */
export function buildSheetImageElement(data: SheetImageData): HTMLTableElement {
  const { sheet, displayValues, rawValues, rowHeights, columnWidths } = data;
  const lastRow = Math.max(0, displayValues.length - 1);
  const lastColumn = Math.max(0, Math.max(0, ...displayValues.map((row) => row.length)) - 1);

  // 合并单元格：记录左上角跨度、标记被覆盖单元格
  const covered = new Set<string>();
  const spans = new Map<string, { rowSpan: number; columnSpan: number }>();
  for (const merge of sheet.mergeData ?? []) {
    const rowSpan = Math.min(merge.endRow, lastRow) - merge.startRow + 1;
    const columnSpan = Math.min(merge.endColumn, lastColumn) - merge.startColumn + 1;
    if (rowSpan <= 0 || columnSpan <= 0) continue;
    spans.set(`${merge.startRow}:${merge.startColumn}`, { rowSpan, columnSpan });
    for (let r = merge.startRow; r < merge.startRow + rowSpan; r++) {
      for (let c = merge.startColumn; c < merge.startColumn + columnSpan; c++) {
        if (r !== merge.startRow || c !== merge.startColumn) covered.add(`${r}:${c}`);
      }
    }
  }

  const table = document.createElement('table');
  const totalWidth = Array.from({ length: lastColumn + 1 }, (_, c) =>
    Math.max(8, columnWidths[c] ?? sheet.columnData?.[c]?.w ?? sheet.defaultColumnWidth ?? DEFAULT_COLUMN_WIDTH),
  ).reduce((sum, w) => sum + w, 0);
  table.setAttribute('cellspacing', '0');
  table.setAttribute('cellpadding', '0');
  table.style.cssText = `border-collapse:collapse;table-layout:fixed;width:${totalWidth}px;background:#fff;font-family:Arial,'PingFang SC','Microsoft YaHei',sans-serif;`;

  const colgroup = document.createElement('colgroup');
  for (let c = 0; c <= lastColumn; c++) {
    const col = document.createElement('col');
    const width = Math.max(
      8,
      columnWidths[c] ?? sheet.columnData?.[c]?.w ?? sheet.defaultColumnWidth ?? DEFAULT_COLUMN_WIDTH,
    );
    col.style.width = `${width}px`;
    colgroup.appendChild(col);
  }
  table.appendChild(colgroup);

  const tbody = document.createElement('tbody');
  const showGridlines = sheet.showGridlines !== 0;
  for (let r = 0; r <= lastRow; r++) {
    const tr = document.createElement('tr');
    const height = Math.max(
      8,
      rowHeights[r] ?? sheet.rowData?.[r]?.h ?? sheet.defaultRowHeight ?? DEFAULT_ROW_HEIGHT,
    );
    tr.style.height = `${height}px`;
    for (let c = 0; c <= lastColumn; c++) {
      const key = `${r}:${c}`;
      if (covered.has(key)) continue;
      const cell = sheet.cellData?.[r]?.[c];
      const style = styleOf(data, cell);
      const span = spans.get(key);
      const raw = rawValues[r]?.[c];
      const text = displayValues[r]?.[c] || richTextOf(cell);
      const td = document.createElement('td');

      if (span) {
        if (span.rowSpan > 1) td.rowSpan = span.rowSpan;
        if (span.columnSpan > 1) td.colSpan = span.columnSpan;
      }

      const css: string[] = [
        'padding:2px 4px',
        'overflow:hidden',
        `white-space:${style?.tb === 3 ? 'pre-wrap' : 'pre'}`,
        `word-break:${style?.tb === 3 ? 'break-word' : 'normal'}`,
        `text-align:${style?.ht === 2 ? 'center' : style?.ht === 3 ? 'right' : defaultAlign(cell, raw)}`,
        `vertical-align:${style?.vt === 1 ? 'top' : style?.vt === 3 ? 'bottom' : 'middle'}`,
        `font-size:${style?.fs ?? 11}pt`,
        'line-height:1.35',
        'color:#1f2329',
      ];
      if (showGridlines) css.push('border:1px solid #e1e3e6');
      if (style?.ff) css.push(`font-family:${style.ff.replace(/["\\;]/g, '')},Arial,'PingFang SC','Microsoft YaHei',sans-serif`);
      if (style?.bl === 1) css.push('font-weight:700');
      if (style?.it === 1) css.push('font-style:italic');
      if (style?.ul?.s === 1) css.push('text-decoration:underline');
      if (style?.st?.s === 1) css.push('text-decoration:line-through');
      const color = resolveFontColor(style);
      if (color) css.push(`color:${color}`);
      const background = resolveBackground(style);
      if (background) css.push(`background:${background}`);
      for (const side of ['t', 'r', 'b', 'l'] as const) {
        const border = borderOf(style, side);
        if (border) css.push(`border-${side === 't' ? 'top' : side === 'r' ? 'right' : side === 'b' ? 'bottom' : 'left'}:${border}`);
      }
      td.style.cssText = css.join(';');
      td.textContent = text;
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  return table;
}

function canvasToBlob(canvas: HTMLCanvasElement): Promise<Blob | null> {
  return new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
}

/** 导出当前 sheet 整表为 PNG（内容超出可视区域也完整保留） */
export async function exportSheetPng(data: SheetImageData, title: string): Promise<void> {
  const host = document.createElement('div');
  host.style.cssText = 'position:fixed;top:0;left:0;z-index:-1;background:#fff;';
  const table = buildSheetImageElement(data);
  host.appendChild(table);
  document.body.appendChild(host);
  try {
    if (document.fonts?.ready) await document.fonts.ready.catch(() => {});
    const width = Math.max(1, table.scrollWidth || table.offsetWidth);
    const height = Math.max(1, table.scrollHeight || table.offsetHeight);
    // 大表限制画布尺寸与总像素：缩放到浏览器可承受范围，避免 toBlob 返回空
    const dpr = typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1;
    const MAX_DIM = 16_000;
    const MAX_PIXELS = 50_000_000;
    const scale = Math.max(
      0.01,
      Math.min(dpr, 2, MAX_DIM / width, MAX_DIM / height, Math.sqrt(MAX_PIXELS / (width * height))),
    );
    const canvas = await html2canvas(host, {
      backgroundColor: '#ffffff',
      logging: false,
      scale,
      width,
      height,
      windowWidth: Math.max(width, window.innerWidth),
      windowHeight: Math.max(height, window.innerHeight),
    });
    const blob = await canvasToBlob(canvas);
    if (blob) downloadBlob(blob, `${title}.png`);
  } finally {
    host.remove();
  }
}
