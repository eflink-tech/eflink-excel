// xlsx 导入预览：用快照里的单元格数据渲染轻量表格（不做公式计算，只展示导入后的值）
import { useMemo, useState } from 'react';
import { useI18n } from '@eflink-tech/office-kit';
import type { CellStyle, SnapshotCell, WorkbookSnapshot } from '../types/spreadsheet';

/** 预览行列上限：覆盖常见首屏，避免大表拖慢弹窗 */
const MAX_ROWS = 20;
const MAX_COLS = 10;

function columnName(index: number): string {
  let name = '';
  let value = index;
  do {
    name = String.fromCharCode(65 + (value % 26)) + name;
    value = Math.floor(value / 26) - 1;
  } while (value >= 0);
  return name;
}

function cellText(cell?: SnapshotCell): string {
  if (!cell) return '';
  const rich = cell.p?.body?.dataStream;
  if (rich) return rich.replace(/[\r\n]+/g, ' ').trim();
  if (cell.v === undefined || cell.v === null || cell.v === '') {
    return cell.f ? `=${cell.f}` : '';
  }
  return String(cell.v);
}

function resolveStyle(snapshot: WorkbookSnapshot, cell: SnapshotCell): CellStyle | undefined {
  if (!cell.s) return undefined;
  return typeof cell.s === 'string' ? snapshot.styles?.[cell.s] : cell.s;
}

function colorValue(color?: { rgb: string }): string | undefined {
  if (!color?.rgb) return undefined;
  return color.rgb;
}

export function XlsxPreview({ snapshot }: { snapshot: WorkbookSnapshot }) {
  const { t } = useI18n();
  const sheetIds = useMemo(
    () =>
      (snapshot.sheetOrder ?? Object.keys(snapshot.sheets)).filter(
        (id) => snapshot.sheets[id] && snapshot.sheets[id].hidden !== 1,
      ),
    [snapshot],
  );
  const [activeId, setActiveId] = useState(sheetIds[0] ?? '');
  const sheet = snapshot.sheets[activeId] ?? snapshot.sheets[sheetIds[0] ?? ''];

  const layout = useMemo(() => {
    if (!sheet) return null;
    let usedRow = -1;
    let usedCol = -1;
    for (const [rowKey, row] of Object.entries(sheet.cellData ?? {})) {
      const rowIndex = Number(rowKey);
      for (const [colKey, cell] of Object.entries(row)) {
        if (!cellText(cell)) continue;
        usedRow = Math.max(usedRow, rowIndex);
        usedCol = Math.max(usedCol, Number(colKey));
      }
    }
    for (const merge of sheet.mergeData ?? []) {
      usedRow = Math.max(usedRow, merge.endRow);
      usedCol = Math.max(usedCol, merge.endColumn);
    }
    const rows = Math.min(MAX_ROWS, Math.max(1, usedRow + 1));
    const cols = Math.min(MAX_COLS, Math.max(1, usedCol + 1));

    const anchors = new Map<string, { rowSpan: number; colSpan: number }>();
    const covered = new Set<string>();
    for (const merge of sheet.mergeData ?? []) {
      if (merge.startRow >= rows || merge.startColumn >= cols) continue;
      const endRow = Math.min(merge.endRow, rows - 1);
      const endCol = Math.min(merge.endColumn, cols - 1);
      anchors.set(`${merge.startRow}:${merge.startColumn}`, {
        rowSpan: endRow - merge.startRow + 1,
        colSpan: endCol - merge.startColumn + 1,
      });
      for (let r = merge.startRow; r <= endRow; r++) {
        for (let c = merge.startColumn; c <= endCol; c++) {
          if (r === merge.startRow && c === merge.startColumn) continue;
          covered.add(`${r}:${c}`);
        }
      }
    }

    const columnWidths = Array.from({ length: cols }, (_, col) => {
      const width = sheet.columnData?.[col]?.w ?? sheet.defaultColumnWidth ?? 88;
      return Math.max(60, Math.min(220, Math.round(width)));
    });
    const rowHeights = Array.from({ length: rows }, (_, row) => {
      const height = sheet.rowData?.[row]?.h ?? sheet.defaultRowHeight ?? 24;
      return Math.max(20, Math.min(48, Math.round(height)));
    });
    return { rows, cols, anchors, covered, columnWidths, rowHeights };
  }, [sheet]);

  if (!sheet || !layout) return null;

  return (
    <div>
      {sheetIds.length > 1 && (
        <div className="mb-2 flex flex-wrap gap-1 border-b border-slate-200 pb-1">
          {sheetIds.map((id) => (
            <button
              key={id}
              type="button"
              className={`rounded-t px-2 py-1 text-xs ${
                id === sheet.id
                  ? 'border-b-2 border-[#107c41] font-medium text-[#107c41]'
                  : 'text-slate-500 hover:bg-slate-100'
              }`}
              onClick={() => setActiveId(id)}
            >
              {snapshot.sheets[id]?.name ?? id}
            </button>
          ))}
        </div>
      )}
      <div className="max-h-[52vh] overflow-auto rounded border border-slate-200">
        <table className="border-collapse bg-white text-[11px] text-slate-700">
          <colgroup>
            <col style={{ width: 36 }} />
            {layout.columnWidths.map((width, col) => (
              <col key={col} style={{ width }} />
            ))}
          </colgroup>
          <thead>
            <tr>
              <th className="sticky left-0 top-0 z-10 bg-slate-100 text-center font-normal text-slate-400" />
              {layout.columnWidths.map((_, col) => (
                <th
                  key={col}
                  className="sticky top-0 z-[1] bg-slate-100 px-1 py-0.5 text-center font-normal text-slate-400"
                >
                  {columnName(col)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: layout.rows }, (_, row) => (
              <tr key={row}>
                <th className="sticky left-0 bg-slate-100 px-1 text-center font-normal text-slate-400">
                  {row + 1}
                </th>
                {Array.from({ length: layout.cols }, (_, col) => {
                  if (layout.covered.has(`${row}:${col}`)) return null;
                  const anchor = layout.anchors.get(`${row}:${col}`);
                  const cell = sheet.cellData?.[row]?.[col];
                  const style = cell ? resolveStyle(snapshot, cell) : undefined;
                  const align = style?.ht === 2 ? 'center' : style?.ht === 3 ? 'right' : 'left';
                  return (
                    <td
                      key={col}
                      rowSpan={anchor?.rowSpan}
                      colSpan={anchor?.colSpan}
                      className="overflow-hidden whitespace-pre text-ellipsis border-b border-r border-slate-200 px-1.5 align-middle"
                      style={{
                        height: layout.rowHeights[row],
                        backgroundColor: colorValue(style?.bg),
                        color: colorValue(style?.cl),
                        fontWeight: style?.bl ? 600 : undefined,
                        fontStyle: style?.it ? 'italic' : undefined,
                        fontSize: style?.fs ? Math.max(10, Math.min(16, style.fs)) : undefined,
                        textAlign: align,
                      }}
                      title={cellText(cell)}
                    >
                      {cellText(cell)}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="mt-2 flex flex-wrap gap-x-4 text-xs text-gray-500">
        <span>{t('excel.import.sheetCount', { count: sheetIds.length })}</span>
        <span>
          {t('excel.import.truncated', { rows: MAX_ROWS, cols: MAX_COLS })}
        </span>
      </div>
    </div>
  );
}
