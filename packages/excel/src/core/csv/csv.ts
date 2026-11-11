// CSV 导入导出（RFC 4180）：纯函数，便于单测；导出带 UTF-8 BOM，避免 Excel 打开中文乱码。
import { createEmptySnapshot, type WorkbookSnapshot } from '../../types/spreadsheet';

/** 解析 CSV 文本为二维字符串。支持 BOM、CRLF/LF、双引号转义与字段内换行 */
export function parseCsv(text: string): string[][] {
  const input = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let fieldStarted = false;

  const pushField = () => {
    row.push(field);
    field = '';
    fieldStarted = false;
  };
  const pushRow = () => {
    pushField();
    rows.push(row);
    row = [];
  };

  for (let i = 0; i < input.length; i++) {
    const ch = input[i];
    if (inQuotes) {
      if (ch === '"') {
        if (input[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }
    if (ch === '"' && !fieldStarted) {
      inQuotes = true;
      fieldStarted = true;
    } else if (ch === ',') {
      pushField();
    } else if (ch === '\r' || ch === '\n') {
      pushRow();
      if (ch === '\r' && input[i + 1] === '\n') i++;
    } else {
      field += ch;
      fieldStarted = true;
    }
  }
  // 文件末尾没有换行时补最后一行；仅有末尾换行则不产生空行
  if (fieldStarted || field.length > 0 || row.length > 0) pushRow();
  return rows;
}

/** 单元格需要加引号：包含分隔符、引号、换行或首尾空白 */
function needsQuotes(value: string): boolean {
  return /[",\r\n]/.test(value) || /^\s|\s$/.test(value);
}

function escapeField(value: string): string {
  return needsQuotes(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

/** 序列化为 CSV 文本：UTF-8 BOM + CRLF，Excel 兼容性最好 */
export function serializeCsv(rows: readonly (readonly string[])[]): string {
  return `\uFEFF${rows.map((row) => row.map(escapeField).join(',')).join('\r\n')}`;
}

/** 去掉末尾全空的行/列，避免把已用区域里的空样式带进文件 */
export function trimEmptyEdges(rows: string[][]): string[][] {
  let lastRow = rows.length - 1;
  while (lastRow >= 0 && rows[lastRow]!.every((value) => value === '')) lastRow--;
  if (lastRow < 0) return [];
  const trimmed = rows.slice(0, lastRow + 1).map((row) => [...row]);
  let lastCol = 0;
  for (const row of trimmed) {
    for (let c = row.length - 1; c >= 0; c--) {
      if (row[c] !== '') {
        lastCol = Math.max(lastCol, c);
        break;
      }
    }
  }
  return trimmed.map((row) => row.slice(0, Math.max(1, lastCol + 1)));
}

/** CSV 文本 → 单表快照（首行不特殊处理，全部按文本导入，保留前导零等原始内容） */
export function csvToSnapshot(text: string, name: string): WorkbookSnapshot {
  const rows = trimEmptyEdges(parseCsv(text));
  const snapshot = createEmptySnapshot(name);
  const sheet = snapshot.sheets['sheet-01']!;
  if (name) sheet.name = name.slice(0, 31);
  const rowCount = Math.max(100, rows.length);
  const columnCount = Math.max(20, ...rows.map((row) => row.length));
  sheet.rowCount = rowCount;
  sheet.columnCount = columnCount;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!;
    for (let c = 0; c < row.length; c++) {
      const value = row[c]!;
      if (value === '') continue;
      (sheet.cellData[r] ??= {})[c] = { v: value, t: 1 };
    }
  }
  return snapshot;
}

/** 快照的当前 sheet 值 → CSV 二维数组（显示值由调用方提供，确保公式结果与数字格式一致） */
export function displayValuesToRows(values: readonly (readonly string[])[]): string[][] {
  return trimEmptyEdges(
    values.map((row) =>
      row.map((value) => {
        if (value == null) return '';
        // 富文本单元格的显示值带段末换行（dataStream 的段落终止符），不是用户内容
        return String(value).replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n+$/, '');
      }),
    ),
  );
}
