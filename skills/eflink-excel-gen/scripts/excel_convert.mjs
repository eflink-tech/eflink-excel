#!/usr/bin/env node
/**
 * SheetSpec → .efx.json（SheetDocument 备份，eflink-excel 编辑器「导入文件」直接打开）
 *            + .snapshot.json（WorkbookSnapshot，模板入库 content 原样使用）
 *            + .html（多工作表预览，供自检）
 * 用法：node excel_convert.mjs my-spec.json -o ./out
 */
import fs from 'node:fs';
import path from 'node:path';

const args = process.argv.slice(2);
const inFile = args.find((a) => !a.startsWith('-'));
const outFlag = args.indexOf('-o');
const outDir = outFlag >= 0 ? args[outFlag + 1] : '.';
if (!inFile) {
  console.error('用法: node excel_convert.mjs my-spec.json -o ./out');
  process.exit(1);
}

const warnings = [];
const warn = (m) => warnings.push(m);

const spec = JSON.parse(fs.readFileSync(inFile, 'utf8'));
if (!spec || typeof spec !== 'object' || Array.isArray(spec)) throw new Error('spec 必须是对象');
if (!Array.isArray(spec.sheets) || !spec.sheets.length) throw new Error('spec.sheets 必须是非空数组');

const BORDER_STYLE = {
  thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7,
  medium: 8, mediumDashed: 9, mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13,
};
const H_ALIGN = { left: 1, center: 2, right: 3 };
const V_ALIGN = { top: 1, middle: 2, bottom: 3 };
const FORMAT_PRESETS = {
  money: '#,##0.00',
  integer: '#,##0',
  percent: '0.0%',
  date: 'yyyy-mm-dd',
  datetime: 'yyyy-mm-dd hh:mm',
};

const hex = (v) => typeof v === 'string' && /^#[0-9a-fA-F]{6}$/.test(v);
const num = (v) => typeof v === 'number' && Number.isFinite(v);

// ── A1 引用 ──
function colToIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
function indexToCol(i) {
  let s = '';
  i += 1;
  while (i > 0) {
    const r = (i - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}
function parseA1(ref) {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(ref.trim());
  return m ? { row: Number(m[2]) - 1, col: colToIndex(m[1]) } : null;
}
function parseMerge(ref) {
  const m = /^\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$/.exec(ref.trim());
  if (!m) return null;
  return { startRow: Number(m[2]) - 1, startColumn: colToIndex(m[1]), endRow: Number(m[4]) - 1, endColumn: colToIndex(m[3]) };
}

// ── 样式映射（高层样式 → Univer CellStyle）──
function toUniverStyle(raw, where) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return undefined;
  const st = {};
  if (raw.bold) st.bl = 1;
  if (raw.italic) st.it = 1;
  if (num(raw.fontSize) && raw.fontSize >= 6 && raw.fontSize <= 96) st.fs = raw.fontSize;
  else if (raw.fontSize !== undefined) warn(`${where}.fontSize 超出 6~96，已忽略`);
  if (typeof raw.font === 'string' && raw.font) st.ff = raw.font;
  if (hex(raw.color)) st.cl = { rgb: raw.color };
  else if (raw.color !== undefined) warn(`${where}.color 非法(${raw.color})，已忽略`);
  if (hex(raw.background)) st.bg = { rgb: raw.background };
  else if (raw.background !== undefined) warn(`${where}.background 非法(${raw.background})，已忽略`);
  if (raw.underline) st.ul = { s: 1 };
  if (raw.strike) st.st = { s: 1 };
  if (raw.align) {
    if (H_ALIGN[raw.align]) st.ht = H_ALIGN[raw.align];
    else warn(`${where}.align="${raw.align}" 非法（left/center/right），已忽略`);
  }
  if (raw.vAlign) {
    if (V_ALIGN[raw.vAlign]) st.vt = V_ALIGN[raw.vAlign];
    else warn(`${where}.vAlign="${raw.vAlign}" 非法（top/middle/bottom），已忽略`);
  }
  if (raw.wrap) st.tb = 3;
  const format = raw.numberFormat ?? raw.format;
  if (format !== undefined) {
    const pattern = FORMAT_PRESETS[format] ?? format;
    if (typeof pattern === 'string' && pattern) st.n = { pattern };
    else warn(`${where}.numberFormat 非法，已忽略`);
  }
  if (num(raw.rotation)) st.tr = { a: raw.rotation, ...(raw.vertical ? { v: 1 } : {}) };
  else if (raw.vertical) st.tr = { a: 0, v: 1 };
  if (raw.border) {
    const conf = typeof raw.border === 'string' ? { style: raw.border } : raw.border === true ? {} : raw.border;
    const styleName = conf.style ?? 'thin';
    const line = BORDER_STYLE[styleName];
    if (!line) warn(`${where}.border.style="${styleName}" 未知，按 thin 处理`);
    const color = hex(conf.color) ? conf.color : '#C9D3E0';
    const side = { s: line || 1, cl: { rgb: color } };
    const sides = Array.isArray(conf.sides) && conf.sides.length ? conf.sides : ['top', 'bottom', 'left', 'right'];
    st.bd = {};
    for (const s of sides) {
      const key = { top: 't', bottom: 'b', left: 'l', right: 'r' }[s];
      if (key) st.bd[key] = { ...side };
    }
  }
  return Object.keys(st).length ? st : undefined;
}

// ── 公式求值（预览用，支持常用函数与四则/比较运算）──
function makeEvaluator(getCell) {
  const memo = new Map();
  const visiting = new Set();
  const numVal = (v) => (typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : v === '' || v == null ? 0 : Number(v) || 0);
  function tokenize(src) {
    const out = [];
    let i = 0;
    while (i < src.length) {
      const ch = src[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (/[0-9.]/.test(ch)) {
        let j = i;
        while (j < src.length && /[0-9.]/.test(src[j])) j++;
        out.push({ t: 'num', v: Number(src.slice(i, j)) });
        i = j;
        continue;
      }
      if (ch === '"') {
        let j = i + 1, s = '';
        while (j < src.length) {
          if (src[j] === '"' && src[j + 1] === '"') { s += '"'; j += 2; continue; }
          if (src[j] === '"') break;
          s += src[j++];
        }
        out.push({ t: 'str', v: s });
        i = j + 1;
        continue;
      }
      const ref = /^\$?[A-Za-z]{1,3}\$?\d+/.exec(src.slice(i));
      if (ref) {
        out.push({ t: 'ref', v: ref[0] });
        i += ref[0].length;
        continue;
      }
      const ident = /^[A-Za-z][A-Za-z0-9.]*/.exec(src.slice(i));
      if (ident) {
        out.push({ t: 'ident', v: ident[0] });
        i += ident[0].length;
        continue;
      }
      const op = /^(<=|>=|<>|[+\-*/^%(),:<>=])/.exec(src.slice(i));
      if (op) {
        out.push({ t: 'op', v: op[0] });
        i += op[0].length;
        continue;
      }
      i++;
    }
    return out;
  }
  function valuesInRange(a, b) {
    const list = [];
    const r1 = Math.min(a.row, b.row), r2 = Math.max(a.row, b.row);
    const c1 = Math.min(a.col, b.col), c2 = Math.max(a.col, b.col);
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) list.push(cellValue(r, c));
    return list;
  }
  function cellValue(r, c) {
    const key = `${r},${c}`;
    if (memo.has(key)) return memo.get(key);
    const raw = getCell(r, c);
    let val = raw?.v;
    if (raw?.f) {
      if (visiting.has(key)) return 0;
      visiting.add(key);
      try { val = evalFormula(raw.f); } catch { val = undefined; } finally { visiting.delete(key); }
    }
    memo.set(key, val);
    return val;
  }
  function evalFormula(formula) {
    const src = formula.replace(/^=/, '');
    const toks = tokenize(src);
    let pos = 0;
    const peek = () => toks[pos];
    const eat = (v) => (peek() && peek().v === v ? (pos++, true) : false);
    function args() {
      const list = [];
      if (peek() && peek().v === ')') return list;
      for (;;) {
        list.push(expr());
        if (!eat(',')) break;
      }
      return list;
    }
    function primary() {
      const tk = toks[pos++];
      if (!tk) return 0;
      if (tk.t === 'num' || tk.t === 'str') return tk.v;
      if (tk.t === 'ref') {
        const a = parseA1(tk.v);
        if (!a) return 0;
        if (peek() && peek().v === ':') {
          pos++;
          const btk = toks[pos++];
          const b = btk?.t === 'ref' ? parseA1(btk.v) : null;
          if (!b) return 0;
          return { range: [a, b], values: valuesInRange(a, b) };
        }
        return cellValue(a.row, a.col);
      }
      if (tk.t === 'ident') {
        const name = tk.v.toUpperCase();
        if (eat('(')) {
          const a = args();
          eat(')');
          return callFn(name, a);
        }
        return 0;
      }
      if (tk.v === '(') {
        const v = expr();
        eat(')');
        return v;
      }
      if (tk.v === '-') return -numVal(primary());
      if (tk.v === '+') return numVal(primary());
      return 0;
    }
    function matchCriteria(v, criteria) {
      const m = typeof criteria === 'string' ? /^(<=|>=|<>|=|<|>)?(.*)$/.exec(criteria) : null;
      const op = m ? m[1] || '=' : '=';
      const target = m ? m[2].trim() : criteria;
      const tNum = Number(target);
      const useNum = target !== '' && !Number.isNaN(tNum);
      const vv = useNum ? Number(v) : String(v ?? '');
      const tt = useNum ? tNum : target;
      if (op === '=') return vv === tt;
      if (op === '<>') return vv !== tt;
      if (op === '>') return vv > tt;
      if (op === '<') return vv < tt;
      if (op === '>=') return vv >= tt;
      if (op === '<=') return vv <= tt;
      return false;
    }
    function callFn(name, argv) {
      const flatten = (v) => (v && typeof v === 'object' && v.values ? v.values : [v]);
      const nums = (list) => list.flatMap(flatten).map(numVal).filter((n) => Number.isFinite(n));
      if (name === 'SUM') return nums(argv).reduce((s, n) => s + n, 0);
      if (name === 'AVERAGE') { const l = nums(argv); return l.length ? l.reduce((s, n) => s + n, 0) / l.length : 0; }
      if (name === 'MIN') { const l = nums(argv); return l.length ? Math.min(...l) : 0; }
      if (name === 'MAX') { const l = nums(argv); return l.length ? Math.max(...l) : 0; }
      if (name === 'COUNT') return nums(argv).length;
      if (name === 'ABS') return Math.abs(numVal(argv[0]));
      if (name === 'ROUND') { const p = Math.pow(10, numVal(argv[1] ?? 0)); return Math.round(numVal(argv[0]) * p) / p; }
      if (name === 'IF') return argv[0] ? argv[1] ?? true : argv[2] ?? false;
      if (name === 'IFERROR') { try { return argv[0]; } catch { return argv[1] ?? ''; } }
      if (name === 'COUNTIF') {
        const range = argv[0]?.values ?? [];
        return range.filter((v) => matchCriteria(v, argv[1])).length;
      }
      if (name === 'SUMIF') {
        const range = argv[0]?.values ?? [];
        const sumRange = argv[2]?.values ?? range;
        return range.reduce((s, v, i) => (matchCriteria(v, argv[1]) ? s + numVal(sumRange[i]) : s), 0);
      }
      if (name === 'AVERAGEIF') {
        const range = argv[0]?.values ?? [];
        const avgRange = argv[2]?.values ?? range;
        const hit = range.reduce((s, v, i) => (matchCriteria(v, argv[1]) ? s + numVal(avgRange[i]) : s), 0);
        const n = range.filter((v) => matchCriteria(v, argv[1])).length;
        return n ? hit / n : 0;
      }
      return 0;
    }
    function power() {
      let a = primary();
      while (peek() && peek().v === '^') { pos++; a = Math.pow(numVal(a), numVal(primary())); }
      return a;
    }
    function term() {
      let a = power();
      while (peek() && (peek().v === '*' || peek().v === '/' || peek().v === '%')) {
        const op = toks[pos++].v;
        const b = power();
        a = op === '*' ? numVal(a) * numVal(b) : op === '/' ? numVal(a) / numVal(b) : numVal(a) / 100;
      }
      return a;
    }
    function expr() {
      let a = term();
      while (peek() && (peek().v === '+' || peek().v === '-')) {
        const op = toks[pos++].v;
        const b = term();
        a = op === '+' ? numVal(a) + numVal(b) : numVal(a) - numVal(b);
      }
      const cmp = ['<=', '>=', '<>', '=', '<', '>'];
      if (peek() && cmp.includes(peek().v)) {
        const op = toks[pos++].v;
        const b = term();
        if (op === '=') a = a === b;
        else if (op === '<>') a = a !== b;
        else if (op === '>') a = numVal(a) > numVal(b);
        else if (op === '<') a = numVal(a) < numVal(b);
        else if (op === '>=') a = numVal(a) >= numVal(b);
        else if (op === '<=') a = numVal(a) <= numVal(b);
      }
      return a;
    }
    return expr();
  }
  return cellValue;
}

function formatNumber(v, pattern) {
  if (typeof v !== 'number' || !Number.isFinite(v) || !pattern) return String(v ?? '');
  if (pattern === '#,##0.00') return v.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  if (pattern === '#,##0') return v.toLocaleString('zh-CN', { maximumFractionDigits: 0 });
  if (pattern === '0.0%') return `${(v * 100).toFixed(1)}%`;
  if (/^0+\.0+%$/.test(pattern)) {
    const digits = (pattern.match(/\.(0+)%/) || [, '0'])[1].length;
    return `${(v * 100).toFixed(digits)}%`;
  }
  return String(v);
}

// ── 构建快照 ──
function buildSheet(rawSheet, idx) {
  const name = typeof rawSheet.name === 'string' && rawSheet.name.trim() ? rawSheet.name.trim() : `Sheet${idx + 1}`;
  const where = `sheets[${idx}]`;
  const data = Array.isArray(rawSheet.data) ? rawSheet.data : [];
  if (!data.length) warn(`${where}.data 为空`);
  const namedStyles = rawSheet.styles && typeof rawSheet.styles === 'object' ? rawSheet.styles : {};
  const usedStyleNames = new Set();
  const resolveStyle = (s, at) => {
    if (s === undefined) return undefined;
    if (typeof s === 'string') {
      if (!namedStyles[s]) { warn(`${at}.style="${s}" 未在 styles 中定义，已忽略`); return undefined; }
      usedStyleNames.add(s);
      return { ...namedStyles[s] };
    }
    return toUniverStyle(s, at);
  };

  const cellData = {};
  let maxRow = -1, maxCol = -1;
  data.forEach((row, r) => {
    if (!Array.isArray(row)) { warn(`${where}.data[${r}] 必须是数组`); return; }
    row.forEach((cell, c) => {
      if (cell === null || cell === undefined || cell === '') return;
      const at = `${where}.data[${r}][${c}]`;
      const obj = typeof cell === 'object' && !Array.isArray(cell) ? cell : { value: cell };
      const out = {};
      if (typeof obj.formula === 'string' && obj.formula) {
        if (!obj.formula.startsWith('=')) warn(`${at}.formula 建议以 = 开头`);
        out.f = obj.formula.startsWith('=') ? obj.formula : `=${obj.formula}`;
      }
      if (obj.value !== undefined) {
        if (typeof obj.value === 'string' && /[\r\n]/.test(obj.value)) {
          warn(`${at}.value 含换行，Univer 普通单元格会丢失换行；建议改用多行文本或拆分单元格`);
        }
        out.v = obj.value;
      }
      if (obj.formula === undefined && out.v === undefined) return;
      const st = resolveStyle(obj.style, at) ?? {};
      if ((obj.format ?? obj.numberFormat) !== undefined) {
        const extra = toUniverStyle({ numberFormat: obj.format ?? obj.numberFormat }, at);
        if (extra) Object.assign(st, extra);
      }
      if (Object.keys(st).length) out.s = st;
      (cellData[r] ??= {})[c] = out;
      maxRow = Math.max(maxRow, r);
      maxCol = Math.max(maxCol, c);
    });
  });
  for (const k of Object.keys(namedStyles)) if (!usedStyleNames.has(k)) warn(`${where}.styles.${k} 未被引用`);

  const merges = [];
  if (rawSheet.merges !== undefined) {
    if (!Array.isArray(rawSheet.merges)) warn(`${where}.merges 必须是 A1 范围字符串数组`);
    else for (const ref of rawSheet.merges) {
      const m = typeof ref === 'string' ? parseMerge(ref) : null;
      if (!m) { warn(`${where}.merges 非法 "${ref}"，已跳过`); continue; }
      if (m.startRow === m.endRow && m.startColumn === m.endColumn) { warn(`${where}.merges "${ref}" 是单个单元格，已跳过`); continue; }
      const hit = merges.find((o) => !(m.endRow < o.startRow || m.startRow > o.endRow || m.endColumn < o.startColumn || m.startColumn > o.endColumn));
      if (hit) { warn(`${where}.merges "${ref}" 与已有合并区重叠，已跳过`); continue; }
      merges.push(m);
    }
  }
  const columns = Array.isArray(rawSheet.columns) ? rawSheet.columns : [];
  const columnData = {};
  columns.forEach((w, i) => {
    if (num(w) && w > 0) columnData[i] = { w };
    else if (w !== undefined && w !== 0) warn(`${where}.columns[${i}] 必须是正数宽度`);
  });
  const rowData = {};
  if (rawSheet.rowHeights && typeof rawSheet.rowHeights === 'object') {
    for (const [k, h] of Object.entries(rawSheet.rowHeights)) {
      const r = Number(k);
      if (Number.isInteger(r) && num(h) && h > 0) rowData[r] = { h };
      else warn(`${where}.rowHeights["${k}"] 非法`);
    }
  }
  const rowCount = Math.max(Number.isInteger(rawSheet.rowCount) ? rawSheet.rowCount : 0, maxRow + 1, 20);
  const columnCount = Math.max(Number.isInteger(rawSheet.columnCount) ? rawSheet.columnCount : 0, maxCol + 1, columns.length, 8);
  let freeze;
  if (rawSheet.freeze && typeof rawSheet.freeze === 'object') {
    const rows = Number.isInteger(rawSheet.freeze.rows) && rawSheet.freeze.rows > 0 ? rawSheet.freeze.rows : 0;
    const cols = Number.isInteger(rawSheet.freeze.columns) && rawSheet.freeze.columns > 0 ? rawSheet.freeze.columns : 0;
    if (rows || cols) freeze = { startRow: rows, startColumn: cols, xSplit: cols, ySplit: rows };
    else warn(`${where}.freeze 行列数都为 0，已忽略`);
  }
  const sheet = {
    id: rawSheet.id ?? `sheet-${String(idx + 1).padStart(2, '0')}`,
    name,
    rowCount,
    columnCount,
    cellData,
    ...(merges.length ? { mergeData: merges } : {}),
    showGridlines: rawSheet.showGridlines === false ? 0 : 1,
    ...(hex(rawSheet.tabColor) ? { tabColor: rawSheet.tabColor } : {}),
    ...(rawSheet.hidden ? { hidden: 1 } : {}),
    ...(num(rawSheet.defaultRowHeight) ? { defaultRowHeight: rawSheet.defaultRowHeight } : {}),
    ...(num(rawSheet.defaultColumnWidth) ? { defaultColumnWidth: rawSheet.defaultColumnWidth } : {}),
    ...(freeze ? { freeze } : {}),
    ...(Object.keys(rowData).length ? { rowData } : {}),
    ...(Object.keys(columnData).length ? { columnData } : {}),
  };
  return { sheet, rawSheet, namedStyles };
}

const built = spec.sheets.map(buildSheet);
const snapshot = {
  id: spec.snapshotId ?? `workbook-${Date.now()}`,
  name: typeof spec.workbookName === 'string' && spec.workbookName ? spec.workbookName : (spec.title ?? '未命名表格'),
  sheetOrder: built.map((b) => b.sheet.id),
  sheets: Object.fromEntries(built.map((b) => [b.sheet.id, b.sheet])),
};
if (new Set(snapshot.sheetOrder).size !== snapshot.sheetOrder.length) throw new Error('sheet id 重复');

const title = typeof spec.title === 'string' && spec.title.trim() ? spec.title.trim() : snapshot.name;
const now = Date.now();
const doc = {
  source: spec.source ?? { app: '易飞办公', url: 'https://eflink.tech/office/excel' },
  id: spec.docId ?? `doc-${now}`,
  title,
  snapshot,
  createdAt: now,
  updatedAt: now,
};

// ── HTML 预览 ──
const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const styleToCss = (raw, namedStyles) => {
  const s = typeof raw === 'string' ? namedStyles[raw] : raw;
  if (!s) return '';
  const css = [];
  if (s.bold) css.push('font-weight:700');
  if (s.italic) css.push('font-style:italic');
  const deco = [s.underline && 'underline', s.strike && 'line-through'].filter(Boolean);
  if (deco.length) css.push(`text-decoration:${deco.join(' ')}`);
  if (num(s.fontSize)) css.push(`font-size:${s.fontSize}px`);
  if (s.font) css.push(`font-family:${s.font}`);
  if (hex(s.color)) css.push(`color:${s.color}`);
  if (hex(s.background)) css.push(`background:${s.background}`);
  if (s.align) css.push(`text-align:${s.align}`);
  if (s.vAlign) css.push(`vertical-align:${s.vAlign}`);
  if (s.wrap) css.push('white-space:pre-wrap');
  if (s.border) {
    const conf = s.border === true ? {} : typeof s.border === 'string' ? { style: s.border } : s.border;
    const style = conf.style ?? 'thin';
    const c = hex(conf.color) ? conf.color : '#C9D3E0';
    const sides = Array.isArray(conf.sides) && conf.sides.length ? conf.sides : ['top', 'bottom', 'left', 'right'];
    const cssStyle = style === 'dashed' || style === 'mediumDashed' ? 'dashed' : style === 'dotted' || style === 'hair' ? 'dotted' : style === 'double' ? 'double' : 'solid';
    const width = /medium/.test(style) ? 2 : /thick/.test(style) ? 3 : 1;
    for (const side of sides) css.push(`border-${side}:${width}px ${cssStyle} ${c}`);
  }
  const rot = s.vertical ? -90 : num(s.rotation) ? s.rotation : 0;
  if (rot) css.push(`transform:rotate(${rot}deg)`);
  return css.join(';');
};

const sheetHtml = built.map(({ sheet, rawSheet, namedStyles }, si) => {
  const evalCell = makeEvaluator((r, c) => sheet.cellData[r]?.[c]);
  const data = rawSheet.data ?? [];
  // 合并覆盖表
  const covered = new Set();
  const spans = {};
  for (const m of sheet.mergeData ?? []) {
    spans[`${m.startRow},${m.startColumn}`] = { rows: m.endRow - m.startRow + 1, cols: m.endColumn - m.startColumn + 1 };
    for (let r = m.startRow; r <= m.endRow; r++) for (let c = m.startColumn; c <= m.endColumn; c++) {
      if (r !== m.startRow || c !== m.startColumn) covered.add(`${r},${c}`);
    }
  }
  const maxRow = Math.max(data.length, ...(sheet.mergeData ?? []).map((m) => m.endRow + 1), 1);
  const maxCol = Math.max(...data.map((r) => (Array.isArray(r) ? r.length : 0)), ...(sheet.mergeData ?? []).map((m) => m.endColumn + 1), 1);
  const cols = sheet.columnData ?? {};
  const rowsH = sheet.rowData ?? {};
  const colgroup = `<colgroup>${Array.from({ length: maxCol }, (_, c) => `<col style="width:${cols[c]?.w ?? sheet.defaultColumnWidth ?? 88}px">`).join('')}</colgroup>`;
  let trs = '';
  for (let r = 0; r < maxRow; r++) {
    let tds = '';
    for (let c = 0; c < maxCol; c++) {
      if (covered.has(`${r},${c}`)) continue;
      const raw = data[r]?.[c];
      const cell = typeof raw === 'object' && raw !== null && !Array.isArray(raw) ? raw : { value: raw };
      const span = spans[`${r},${c}`];
      const cs = span?.cols > 1 ? ` colspan="${span.cols}"` : '';
      const rs = span?.rows > 1 ? ` rowspan="${span.rows}"` : '';
      const value = cell?.value !== undefined ? cell.value : cell?.formula ? evalCell(r, c) : '';
      const highStyle = typeof cell?.style === 'string' ? namedStyles[cell.style] : cell?.style;
      const patternRaw = highStyle?.numberFormat ?? highStyle?.format ?? cell?.format;
      const pattern = FORMAT_PRESETS[patternRaw] ?? patternRaw;
      const display = typeof value === 'number' ? formatNumber(value, pattern) : typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value ?? '';
      const css = [styleToCss(cell?.style, namedStyles), 'padding:2px 6px;overflow:hidden'].filter(Boolean).join(';');
      tds += `<td${cs}${rs} style="${css}">${esc(display)}</td>`;
    }
    trs += `<tr style="height:${rowsH[r]?.h ?? sheet.defaultRowHeight ?? 24}px">${tds}</tr>`;
  }
  const active = si === 0 ? 'active' : '';
  const freezeParts = [];
  if (sheet.freeze?.ySplit) freezeParts.push(`前 ${sheet.freeze.ySplit} 行`);
  if (sheet.freeze?.xSplit) freezeParts.push(`前 ${sheet.freeze.xSplit} 列`);
  const freezeNote = freezeParts.length ? `<div class="freeze">冻结：${freezeParts.join(' / ')}</div>` : '';
  return `<section class="sheet ${active}" data-idx="${si}">
<div class="bar"><strong>${esc(sheet.name)}</strong>${sheet.hidden ? '<span class="badge">隐藏</span>' : ''}${freezeNote}</div>
<div class="scroll"><table style="border-collapse:separate;border-spacing:0;table-layout:fixed;${sheet.showGridlines === 0 ? '' : 'background-image:linear-gradient(#E5E7EB 1px,transparent 1px),linear-gradient(90deg,#E5E7EB 1px,transparent 1px);background-size:100% 24px,88px 100%;'}">${colgroup}${trs}</table></div>
</section>`;
}).join('\n');

const tabs = built.map(({ sheet }, si) => `<button class="tab ${si === 0 ? 'active' : ''}" data-idx="${si}">${esc(sheet.name)}</button>`).join('');
const html = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>${esc(title)}</title><style>
body{margin:0;background:#F3F4F6;font-family:'Microsoft YaHei',sans-serif;color:#111}
.wrap{max-width:1280px;margin:18px auto;background:#fff;border:1px solid #D1D5DB;border-radius:6px;overflow:hidden}
.tabs{display:flex;gap:2px;padding:8px 10px 0;background:#F9FAFB;border-bottom:1px solid #E5E7EB}
.tab{border:1px solid #D1D5DB;border-bottom:none;background:#F3F4F6;padding:6px 16px;border-radius:4px 4px 0 0;cursor:pointer;font-size:13px}
.tab.active{background:#fff;font-weight:700;color:#1F3864}
.sheet{display:none;padding:12px}
.sheet.active{display:block}
.bar{display:flex;align-items:center;gap:10px;margin-bottom:8px;font-size:13px;color:#374151}
.badge{background:#FEF3C7;color:#92400E;border-radius:3px;padding:1px 6px;font-size:12px}
.freeze{color:#6B7280;font-size:12px}
.scroll{overflow:auto;max-height:80vh;border:1px solid #E5E7EB}
td{white-space:nowrap;font-size:13px;line-height:1.4;box-sizing:border-box}
</style></head><body><div class="wrap"><div class="tabs">${tabs}</div>${sheetHtml}</div>
<script>function show(i){document.querySelectorAll('.tab').forEach(t=>t.classList.toggle('active',Number(t.dataset.idx)===i));document.querySelectorAll('.sheet').forEach(s=>s.classList.toggle('active',Number(s.dataset.idx)===i));}
document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>show(Number(b.dataset.idx)));
show(Number(new URLSearchParams(location.search).get('sheet')||location.hash.slice(1)||0));</script>
</body></html>`;

// ── 写文件 ──
const base = path.basename(inFile).replace(/\.json$/i, '');
fs.mkdirSync(outDir, { recursive: true });
const efxPath = path.join(outDir, `${base}.efx.json`);
const snapPath = path.join(outDir, `${base}.snapshot.json`);
const htmlPath = path.join(outDir, `${base}.html`);
fs.writeFileSync(efxPath, JSON.stringify(doc, null, 2));
fs.writeFileSync(snapPath, JSON.stringify(snapshot, null, 2));
fs.writeFileSync(htmlPath, html);

console.log(`OK ${base}: ${built.length} 个工作表，${built.reduce((s, b) => s + Object.values(b.sheet.cellData).reduce((t, row) => t + Object.keys(row).length, 0), 0)} 个单元格`);
for (const { sheet } of built) console.log(`  - ${sheet.name}: ${sheet.rowCount} 行 × ${sheet.columnCount} 列${sheet.freeze ? `，冻结 ${sheet.freeze.ySplit} 行 / ${sheet.freeze.xSplit} 列` : ''}`);
console.log(`  ${efxPath}`);
console.log(`  ${snapPath}`);
console.log(`  ${htmlPath}`);
if (warnings.length) {
  console.error(`⚠ ${warnings.length} 条警告:`);
  for (const w of warnings) console.error(`  - ${w}`);
}
