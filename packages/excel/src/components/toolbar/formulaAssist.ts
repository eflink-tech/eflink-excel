/** 公式栏补全/引用高亮的纯逻辑：文本扫描、函数签名、A1 引用解析 */

export interface SuggestItem {
  name: string;
  desc: string;
}

/** 公式引用高亮样式：橙色虚线 + 浅橙底 */
export const REF_STROKE = '#f97316';
export const REF_STROKE_DASH = 8;
export const REF_FILL_ALPHA = 0.1;

const TOKEN_CHAR = /[A-Za-z0-9._]/;

/** 当前光标所在"单词"（函数名前缀），跳过字符串与工作表名引号 */
export function readTokenAtCaret(text: string, caret: number): { token: string; start: number; inString: boolean } {
  const end = Math.max(0, Math.min(caret, text.length));
  let inString = false;
  let inSheetQuote = false;
  for (let i = 0; i < end; i++) {
    const ch = text[i];
    if (inString) {
      if (ch === '"') {
        if (text[i + 1] === '"' && i + 1 < end) i++;
        else inString = false;
      }
      continue;
    }
    if (ch === '"') inString = true;
    else if (!inSheetQuote && ch === "'") inSheetQuote = true;
    else if (inSheetQuote && ch === "'") inSheetQuote = false;
  }
  let start = end;
  while (start > 0 && TOKEN_CHAR.test(text[start - 1])) start--;
  if (start > 0 && (text[start - 1] === '$' || text[start - 1] === '"')) start = end;
  return { token: text.slice(start, end), start, inString };
}

/** 光标处若在函数参数内，返回最内层函数名与当前参数序号 */
export function functionContextAtCaret(text: string, caret: number): { name: string; argIndex: number } | null {
  const end = Math.max(0, Math.min(caret, text.length));
  const stack: { name: string; commas: number }[] = [];
  let inString = false;
  let inSheetQuote = false;
  for (let i = 0; i < end; i++) {
    const ch = text[i];
    if (inString) {
      if (ch === '"') {
        if (text[i + 1] === '"' && i + 1 < end) i++;
        else inString = false;
      }
      continue;
    }
    if (inSheetQuote) {
      if (ch === "'") inSheetQuote = false;
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === "'") inSheetQuote = true;
    else if (ch === '(') stack.push({ name: readFunctionName(text, i), commas: 0 });
    else if (ch === ')') stack.pop();
    else if (ch === ',' && stack.length > 0) stack[stack.length - 1].commas++;
  }
  const top = stack[stack.length - 1];
  if (!top || !top.name) return null;
  return { name: top.name, argIndex: top.commas };
}

function readFunctionName(text: string, parenIndex: number): string {
  let start = parenIndex;
  while (start > 0 && /[A-Za-z0-9._]/.test(text[start - 1])) start--;
  return text.slice(start, parenIndex);
}

/** 把函数名补全进公式，已跟"("时不再补括号 */
export function applyFunctionName(text: string, caret: number, name: string): { text: string; caret: number } {
  const { token, start } = readTokenAtCaret(text, caret);
  if (!token) return { text, caret };
  const end = start + token.length;
  const hasParen = /^\s*\(/.test(text.slice(end));
  const insert = hasParen ? name : `${name}(`;
  return {
    text: text.slice(0, start) + insert + text.slice(end),
    caret: start + insert.length,
  };
}

/** 前缀匹配函数表，按字母序、完全匹配优先 */
export function filterSuggestions(
  list: SuggestItem[],
  token: string,
  limit = 10,
): SuggestItem[] {
  const prefix = token.toUpperCase();
  if (!prefix) return [];
  const items = list.filter((item) => item.name.toUpperCase().startsWith(prefix));
  items.sort((a, b) => {
    const aExact = a.name.toUpperCase() === prefix ? 0 : 1;
    const bExact = b.name.toUpperCase() === prefix ? 0 : 1;
    return aExact - bExact || a.name.localeCompare(b.name);
  });
  return items.slice(0, limit);
}

export interface SignatureParam {
  label: string;
  active: boolean;
}

/** 组装 Excel 式签名：可选参数加方括号，循环参数尾部补 … */
export function buildSignature(
  params: { name: string; require?: boolean; repeat?: boolean }[],
  argIndex: number,
): SignatureParam[] {
  if (params.length === 0) return [];
  const active = Math.max(0, Math.min(argIndex, params.length - 1));
  return params.map((param, index) => {
    let label = param.require === false ? `[${param.name}]` : param.name;
    if (param.repeat && index === params.length - 1) label += ', ...';
    return { label, active: index === active };
  });
}

/** 屏蔽双引号字符串，避免把 "A1" 之类的文本当成引用高亮 */
function maskStrings(text: string): string {
  let out = '';
  let inString = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') {
      if (inString && text[i + 1] === '"') {
        out += '  ';
        i++;
        continue;
      }
      inString = !inString;
      out += '"';
      continue;
    }
    out += inString ? ' ' : ch;
  }
  return out;
}

const A1_BODY = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/;
const COL_RANGE_BODY = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/;
const ROW_RANGE_BODY = /^\$?(\d+):\$?(\d+)$/;
const REF_SHAPE = /(?:'([^']*(?:''[^']*)*)'|([A-Za-z0-9_.]+))!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+/g;

export interface RefToken {
  text: string;
  sheetName: string | null;
}

function columnIndex(letters: string): number | null {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n >= 1 && n <= 16384 ? n : null;
}

export interface ParsedRef {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

/** 解析 A1 / A1:B2 / A:C / 1:3，行列均为 0 基，column 上限 16384、row 上限 1048576 */
export function parseA1Body(body: string): ParsedRef | null {
  const a1 = A1_BODY.exec(body);
  if (a1) {
    const sc = columnIndex(a1[1]);
    const sr = Number(a1[2]);
    const ec = a1[3] ? columnIndex(a1[3]) : sc;
    const er = a1[4] ? Number(a1[4]) : sr;
    if (sc == null || ec == null || sr < 1 || er < 1 || sr > 1048576 || er > 1048576) return null;
    return {
      startRow: Math.min(sr, er) - 1,
      startColumn: Math.min(sc, ec) - 1,
      endRow: Math.max(sr, er) - 1,
      endColumn: Math.max(sc, ec) - 1,
    };
  }
  const cols = COL_RANGE_BODY.exec(body);
  if (cols) {
    const sc = columnIndex(cols[1]);
    const ec = columnIndex(cols[2]);
    if (sc == null || ec == null) return null;
    return {
      startRow: 0,
      startColumn: Math.min(sc, ec) - 1,
      endRow: 1048575,
      endColumn: Math.max(sc, ec) - 1,
    };
  }
  const rows = ROW_RANGE_BODY.exec(body);
  if (rows) {
    const sr = Number(rows[1]);
    const er = Number(rows[2]);
    if (sr < 1 || er < 1 || sr > 1048576 || er > 1048576) return null;
    return {
      startRow: Math.min(sr, er) - 1,
      startColumn: 0,
      endRow: Math.max(sr, er) - 1,
      endColumn: 16383,
    };
  }
  return null;
}

/** 从公式文本中提取引用（跳过字符串、函数调用名与非法引用） */
export function extractRefs(text: string): RefToken[] {
  const masked = maskStrings(text);
  const refs: RefToken[] = [];
  REF_SHAPE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = REF_SHAPE.exec(masked))) {
    const start = match.index;
    const raw = match[0];
    const before = start > 0 ? masked[start - 1] : '';
    if (before && /[A-Za-z0-9_.]/.test(before)) continue;
    // 函数名后紧跟 "(" 时不是引用（如 LOG10、BIN2DEC 等带数字的函数名）
    const after = masked[start + raw.length] ?? '';
    if (/[A-Za-z0-9_(]/.test(after)) continue;
    const quotedSheet = match[1];
    const plainSheet = match[2];
    const hasSheet = quotedSheet !== undefined || plainSheet !== undefined;
    const body = hasSheet ? raw.slice(raw.lastIndexOf('!') + 1) : raw;
    if (!parseA1Body(body)) continue;
    const sheetName = quotedSheet !== undefined
      ? quotedSheet.replace(/''/g, "'")
      : plainSheet ?? null;
    refs.push({ text: raw, sheetName });
  }
  return refs;
}
