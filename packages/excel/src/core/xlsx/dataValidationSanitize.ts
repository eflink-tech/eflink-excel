// WPS 常把数据验证写成整列（如 A3:A1048576）。ExcelJS 解析时会把区域逐单元格展开，
// 上百万次 Map 写入会让浏览器导入直接假死；快照又不保留数据验证，因此导入前把超大区域
// 收拢到工作表已用范围。正常大小的验证保持原样。
import JSZip from 'jszip';

const SHEET_ENTRY = /^xl\/worksheets\/[^/]+\.xml$/;
const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;
/** 超过该单元格数的验证区域才收拢，避免改变普通文件的行为 */
const EXPAND_LIMIT = 20_000;

interface Rect {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

const CELL_REF = /^([A-Z]+)(\d+)$/;

function columnNumber(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i += 1) n = n * 26 + letters.charCodeAt(i) - 64;
  return n;
}

function columnLetters(num: number): string {
  let out = '';
  for (let n = num; n > 0; n = Math.floor((n - 1) / 26)) {
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
  }
  return out;
}

function parseCell(ref: string): { row: number; col: number } | null {
  const m = CELL_REF.exec(ref.trim());
  return m ? { row: Number(m[2]), col: columnNumber(m[1]) } : null;
}

function encodeCell(row: number, col: number): string {
  return `${columnLetters(col)}${row}`;
}

/** "A1" / "A1:C9" → 规范化矩形；无法解析返回 null */
function parseArea(area: string): Rect | null {
  const parts = area.split(':');
  const start = parseCell(parts[0]);
  if (!start) return null;
  if (parts.length === 1) return { top: start.row, left: start.col, bottom: start.row, right: start.col };
  if (parts.length !== 2) return null;
  const end = parseCell(parts[1]);
  if (!end) return null;
  return {
    top: Math.min(start.row, end.row),
    left: Math.min(start.col, end.col),
    bottom: Math.max(start.row, end.row),
    right: Math.max(start.col, end.col),
  };
}

function usedRectOf(xml: string): Rect | null {
  const dim = /<dimension\b[^>]*\bref="([^"]+)"/.exec(xml)?.[1];
  if (dim) {
    const rect = parseArea(dim);
    if (rect) return rect;
  }
  let bottom = 0;
  let right = 0;
  for (const m of xml.matchAll(/<c\b[^>]*\br="([A-Z]+\d+)"/g)) {
    const cell = parseCell(m[1]);
    if (!cell) continue;
    bottom = Math.max(bottom, cell.row);
    right = Math.max(right, cell.col);
  }
  return bottom && right ? { top: 1, left: 1, bottom, right } : null;
}

function clampArea(area: string, used: Rect): string | null {
  const rect = parseArea(area);
  if (!rect) return area;
  const rows = rect.bottom - rect.top + 1;
  const cols = rect.right - rect.left + 1;
  if (rows * cols <= EXPAND_LIMIT && rect.bottom <= MAX_ROW && rect.right <= MAX_COL) return area;

  const top = Math.max(rect.top, 1);
  const left = Math.max(rect.left, 1);
  const bottom = Math.min(rect.bottom, used.bottom, MAX_ROW);
  if (top > bottom) return null;
  // 先收行；整行式区域（A1:XFD…）收完仍过大时再收列，避免误改窄列验证
  const right = (bottom - top + 1) * cols <= EXPAND_LIMIT ? rect.right : Math.min(rect.right, Math.max(used.right, left), MAX_COL);
  if (left > right) return null;
  return area.includes(':') ? `${encodeCell(top, left)}:${encodeCell(bottom, right)}` : encodeCell(top, left);
}

function clampSheetXml(xml: string, used: Rect): string {
  return xml.replace(
    /<dataValidation\b[^>]*\/>|<dataValidation\b[^>]*>[\s\S]*?<\/dataValidation>/g,
    (element) => {
      const tagEnd = element.indexOf('>') + 1;
      const tag = element.slice(0, tagEnd);
      const sqref = /\bsqref="([^"]*)"/.exec(tag);
      if (!sqref) return element;

      const areas = sqref[1].split(/\s+/).filter(Boolean);
      const kept = areas.map((area) => clampArea(area, used)).filter((area): area is string => area !== null);
      if (kept.length === areas.length && kept.every((area, i) => area === areas[i])) return element;
      if (!kept.length) return '';
      return tag.replace(sqref[0], `sqref="${kept.join(' ')}"`) + element.slice(tagEnd);
    },
  );
}

/**
 * 收敛工作簿中会触发 ExcelJS 逐单元格展开的超大数据验证区域。
 * 无需修改时直接返回原 buffer，避免额外的解压/压缩开销。
 */
export async function clampOversizedDataValidations(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer);
  const sheets = Object.keys(zip.files).filter((name) => SHEET_ENTRY.test(name));
  let changed = false;

  for (const name of sheets) {
    const entry = zip.file(name);
    if (!entry) continue;
    const xml = await entry.async('string');
    if (!xml.includes('<dataValidation')) continue;
    const used = usedRectOf(xml);
    if (!used) continue;
    const next = clampSheetXml(xml, used);
    if (next !== xml) {
      zip.file(name, next);
      changed = true;
    }
  }

  return changed ? zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' }) : buffer;
}
