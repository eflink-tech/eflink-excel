# xlsx 导入导出增强 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为 `@eflink-tech/excel` 打通 xlsx 导入导出全链路：菜单挂载、替换当前文档式导入、样式/冻结/隐藏/富文本的尽可能完整双向兼容。

**Architecture:** 沿用现有 exceljs 双向映射层（`src/core/xlsx/`），把 `CellStyle`/`SnapshotSheet` 向 Univer 快照形状扩展（快照本就 `as never` 直通 `createWorkbook`，持久化链路零改动）；导入以「替换当前文档 + reloadToken 重载」接入 UI，不新增第三方依赖。

**Tech Stack:** exceljs 4.4（文件解析/生成）、Univer 0.25（渲染引擎，样式键名直接对齐其 `IStyleData`）、Vitest 4 + jsdom。

**工作目录约定：** 下文所有相对路径基于 `eflink-excel/packages/excel/`。测试命令在该目录执行：
`npx vitest run src/core/xlsx`（单模块）或 `npx vitest run`（全量）。

**对应设计文档:** `eflink-excel/docs/superpowers/specs/2026-09-17-xlsx-import-export-design.md`

---

### Task 1: 类型扩展（types/spreadsheet.ts）

**Files:**
- Modify: `src/types/spreadsheet.ts`

- [ ] **Step 1.1: 扩展 CellStyle 并新增富文本/边框类型**

在 `CellStyle` 定义处（文件顶部，`SnapshotCell` 之前）替换为：

```ts
/** 四边边框的单边样式；s 取 Univer BorderStyleTypes 数值枚举（1=THIN … 13=THICK），cl 为边框色 */
export interface BorderStyle {
  s: number;
  cl: { rgb: string };
}

/** 富文本 run：区间 [st, ed) 的样式（形状对齐 Univer ITextRun） */
export interface SheetRichTextRun {
  st: number;
  ed: number;
  ts?: CellStyle;
}

/** 富文本正文子集（形状对齐 Univer IDocumentData.body） */
export interface SheetRichTextBody {
  dataStream: string;
  textRuns?: SheetRichTextRun[];
}

/** 单元格富文本子集（形状对齐 Univer IDocumentData；快照经 as never 直通 createWorkbook） */
export interface SheetRichText {
  id: string;
  body: SheetRichTextBody;
  documentStyle: Record<string, unknown>;
}

export interface CellStyle {
  bl?: 0 | 1;
  it?: 0 | 1;
  fs?: number;
  cl?: { rgb: string };
  bg?: { rgb: string };
  n?: { pattern: string };
  /** 字体族（如「微软雅黑」「Arial」） */
  ff?: string;
  /** 下划线（形状对齐 Univer ITextDecoration 的 show 开关） */
  ul?: { s: 0 | 1 };
  /** 删除线 */
  st?: { s: 0 | 1 };
  /** 水平对齐：Univer HorizontalAlign（1=左 2=中 3=右） */
  ht?: number;
  /** 垂直对齐：Univer VerticalAlign（1=上 2=中 3=下） */
  vt?: number;
  /** 换行策略：Univer WrapStrategy（3=自动换行） */
  tb?: number;
  /** 四边边框 */
  bd?: { t?: BorderStyle; b?: BorderStyle; l?: BorderStyle; r?: BorderStyle };
}

export interface SnapshotCell {
  v?: string | number | boolean;
  f?: string;
  /** 富文本（Univer 文档子集；xlsx 导入方向产出，渲染/持久化由 Univer 原生支持） */
  p?: SheetRichText;
  /** 内联样式对象（导入方向）或 Univer 归一化样式 id（运行时快照，经 workbook.styles 解析） */
  s?: CellStyle | string;
}
```

- [ ] **Step 1.2: 扩展 SnapshotSheet 的 freeze 与隐藏行列标记**

`SnapshotSheet` 中把 `rowData`/`columnData` 两行替换，并新增 `freeze`：

```ts
  /** 冻结窗格：startRow/startColumn = 冻结区行/列数（与 Univer IFreeze 的 yAxisSplit/xAxisSplit 语义一致） */
  freeze?: { startRow: number; startColumn: number; xAxisSplit: number; yAxisSplit: number };
  /** h: 行高(px)；hd: 1 表示隐藏 */
  rowData?: Record<number, { h?: number; hd?: 0 | 1 }>;
  /** w: 列宽(px)；hd: 1 表示隐藏 */
  columnData?: Record<number, { w?: number; hd?: 0 | 1 }>;
```

- [ ] **Step 1.3: 类型检查通过**

Run: `npx tsc -b --noEmit 2>&1 | head -20`（在 packages/excel 下；或 `pnpm --filter @eflink-tech/excel typecheck`）
Expected: 无类型错误（`rowData.h` 变可选只放宽不收紧，现有代码 `if (rd.h)` 兼容）

- [ ] **Step 1.4: Commit**

```bash
git add src/types/spreadsheet.ts
git commit -m "feat: CellStyle/SnapshotSheet 扩展边框/对齐/冻结/隐藏/富文本类型"
```

---

### Task 2: 导入方向 — 单元格样式（边框/对齐/换行/字体族/下划线/删除线）

**Files:**
- Modify: `src/core/xlsx/workbookToSnapshot.ts`
- Test: `src/core/xlsx/xlsx.test.ts`

- [ ] **Step 2.1: 写失败测试**

在 `xlsx.test.ts` 末尾（describe 之外新增 describe）追加：

```ts
import type { CellStyle } from '../../types/spreadsheet';

describe('xlsx 导入方向：扩展样式', () => {
  function importWorkbook(build: (ws: ExcelJS.Worksheet) => void) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('样式表');
    build(ws);
    return workbookToSnapshot(wb, '样式测试');
  }

  it('字体族/下划线/删除线/对齐/换行导入为 Univer 形状样式', () => {
    const snap = importWorkbook((ws) => {
      const a1 = ws.getCell('A1');
      a1.value = '混合样式';
      a1.font = { name: '微软雅黑', underline: true, strike: true, size: 12 };
      a1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    const st = snap.sheets['sheet-01'].cellData[0]?.[0]?.s as CellStyle;
    expect(st.ff).toBe('微软雅黑');
    expect(st.ul).toEqual({ s: 1 });
    expect(st.st).toEqual({ s: 1 });
    expect(st.ht).toBe(2);
    expect(st.vt).toBe(2);
    expect(st.tb).toBe(3);
  });

  it('四边边框（线型枚举 + 颜色）导入，缺省颜色补黑色', () => {
    const snap = importWorkbook((ws) => {
      const a1 = ws.getCell('A1');
      a1.value = '边框';
      a1.border = {
        top: { style: 'thin', color: { argb: 'FFFF0000' } },
        bottom: { style: 'double', color: { argb: 'FF00FF00' } },
        left: { style: 'dashed' },
        right: { style: 'thick', color: { argb: 'FF0000FF' } },
      };
    });
    const st = snap.sheets['sheet-01'].cellData[0]?.[0]?.s as CellStyle;
    expect(st.bd?.t).toEqual({ s: 1, cl: { rgb: '#ff0000' } });
    expect(st.bd?.b).toEqual({ s: 7, cl: { rgb: '#00ff00' } });
    expect(st.bd?.l).toEqual({ s: 4, cl: { rgb: '#000000' } });
    expect(st.bd?.r).toEqual({ s: 13, cl: { rgb: '#0000ff' } });
  });

  it('普通单元格不产出空样式', () => {
    const snap = importWorkbook((ws) => {
      ws.getCell('A1').value = '普通';
    });
    expect(snap.sheets['sheet-01'].cellData[0]?.[0]?.s).toBeUndefined();
  });
});
```

- [ ] **Step 2.2: 跑测试确认失败**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 新增 3 个用例 FAIL（`ff`/`ul`/`bd` 等 undefined），既有用例仍 PASS

- [ ] **Step 2.3: 实现 readStyle 扩展**

`workbookToSnapshot.ts` 顶部类型导入行改为：

```ts
import type { Cell, CellErrorValue, CellFormulaValue, CellRichTextValue, Workbook, Worksheet } from 'exceljs';
import type { BorderStyle, CellStyle, MergeRange, SnapshotCell, SnapshotSheet, WorkbookSnapshot } from '../../types/spreadsheet';
```

在 `CHAR_TO_PX` 常量后追加映射表与边框读取：

```ts
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

function readBorderSide(side: BorderSide): BorderStyle | undefined {
  const s = side?.style ? BORDER_STYLE_IDS[side.style] : undefined;
  if (!s) return undefined;
  return { s, cl: { rgb: side?.color?.argb ? fromArgb(side.color.argb) : '#000000' } };
}
```

`readStyle` 整体替换为：

```ts
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
```

- [ ] **Step 2.4: 跑测试确认通过**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 全部 PASS（含既有用例）

- [ ] **Step 2.5: Commit**

```bash
git add src/core/xlsx/workbookToSnapshot.ts src/core/xlsx/xlsx.test.ts
git commit -m "feat: xlsx 导入支持边框/对齐/换行/字体族/下划线/删除线"
```

---

### Task 3: 导入方向 — 冻结窗格与隐藏行列

**Files:**
- Modify: `src/core/xlsx/workbookToSnapshot.ts`
- Test: `src/core/xlsx/xlsx.test.ts`

- [ ] **Step 3.1: 写失败测试**

在 Task 2 的 describe 内追加：

```ts
  it('冻结窗格与隐藏行列导入', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('冻结表');
    ws.getCell('A1').value = '冻结区';
    ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 2, topLeftCell: 'B3' }] as never;
    ws.getRow(4).hidden = true;
    ws.getColumn(3).hidden = true;
    const snap = workbookToSnapshot(wb, '冻结测试');
    const sheet = snap.sheets['sheet-01'];
    expect(sheet.freeze).toEqual({ startRow: 2, startColumn: 1, xAxisSplit: 1, yAxisSplit: 2 });
    expect(sheet.rowData?.[3]).toEqual({ h: 0, hd: 1 });
    expect(sheet.columnData?.[2]).toEqual({ w: 0, hd: 1 });
  });

  it('非冻结视图不产出 freeze', () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('普通表');
    const snap = workbookToSnapshot(wb, '普通');
    expect(snap.sheets['sheet-01'].freeze).toBeUndefined();
  });
```

- [ ] **Step 3.2: 跑测试确认失败**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 2 个新用例 FAIL（freeze undefined、rowData 无 hd）

- [ ] **Step 3.3: 实现 readFreeze 与隐藏读取**

`readSheet` 中 `rowData`/`columnData` 构建段替换为：

```ts
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
```

`readSheet` 返回对象中新增 `freeze: readFreeze(ws),`（放在 `mergeData` 之后），并在文件内新增：

```ts
/** exceljs 冻结视图（xSplit/ySplit=冻结行列数）→ Univer freeze（startRow/startColumn=滚动区起点） */
function readFreeze(ws: Worksheet): SnapshotSheet['freeze'] {
  const view = ws.views?.[0] as { state?: string; xSplit?: number; ySplit?: number } | undefined;
  if (!view || view.state !== 'frozen') return undefined;
  const x = view.xSplit ?? 0;
  const y = view.ySplit ?? 0;
  if (!x && !y) return undefined;
  return { startRow: y, startColumn: x, xAxisSplit: x, yAxisSplit: y };
}
```

- [ ] **Step 3.4: 跑测试确认通过**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 全部 PASS

- [ ] **Step 3.5: Commit**

```bash
git add src/core/xlsx/workbookToSnapshot.ts src/core/xlsx/xlsx.test.ts
git commit -m "feat: xlsx 导入支持冻结窗格与隐藏行列"
```

---

### Task 4: 导入方向 — 富文本 → cell.p

**Files:**
- Modify: `src/core/xlsx/workbookToSnapshot.ts`
- Test: `src/core/xlsx/xlsx.test.ts`

- [ ] **Step 4.1: 写失败测试**

在 Task 2 的 describe 内追加：

```ts
  it('富文本单元格导入为 cell.p（Univer 文档子集）', () => {
    const snap = importWorkbook((ws) => {
      ws.getCell('A1').value = {
        richText: [
          { text: '普通', font: { size: 12 } },
          { text: '红色', font: { bold: true, color: { argb: 'FFFF0000' } } },
          { text: '结尾' },
        ],
      };
    });
    const cell = snap.sheets['sheet-01'].cellData[0]?.[0];
    expect(cell?.p?.body.dataStream).toBe('普通红色结尾\r\n');
    const runs = cell?.p?.body.textRuns ?? [];
    expect(runs).toHaveLength(1); // 无样式 run 不产出
    expect(runs[0]).toMatchObject({ st: 2, ed: 4, ts: { bl: 1, cl: { rgb: '#ff0000' } } });
  });

  it('富文本含换行时按展开后的流计算 run 区间', () => {
    const snap = importWorkbook((ws) => {
      ws.getCell('A1').value = {
        richText: [
          { text: 'A\nB', font: { bold: true } },
          { text: 'C' },
        ],
      };
    });
    const cell = snap.sheets['sheet-01'].cellData[0]?.[0];
    // A\n 展开为 A\r\n（3 字符），C 在流中位于 [3, 4)
    expect(cell?.p?.body.dataStream).toBe('A\r\nBC\r\n');
    expect(cell?.p?.body.textRuns?.[0]?.ts?.bl).toBe(1);
  });
```

- [ ] **Step 4.2: 跑测试确认失败**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 2 个新用例 FAIL（`p` 为 undefined，现状降级纯文本）

- [ ] **Step 4.3: 实现 rich text 读取**

类型导入行加入 `SheetRichTextRun`：

```ts
import type { BorderStyle, CellStyle, MergeRange, SheetRichTextRun, SnapshotCell, SnapshotSheet, WorkbookSnapshot } from '../../types/spreadsheet';
```

`readCell` 中富文本分支替换（原「富文本第一版降级为纯文本」分支）：

```ts
  } else if (val != null && typeof val === 'object' && 'richText' in (val as CellRichTextValue)) {
    applyRichText(uc, val as CellRichTextValue);
  } else if (val instanceof Date) {
```

文件末尾新增（exceljs 的 `t.font` 字段名与 CellStyle 有差异（`name`/`color.argb`/`underline` 布尔），先做一层显式转换）：

```ts
const RICH_TEXT_DOC_ID = '__eflink-rich-text';

/** exceljs 富文本 → Univer 文档子集：拼接 dataStream（\r\n 为段落分隔，流以 \r\n 结尾）+ 按区间切 textRuns */
function applyRichText(uc: SnapshotCell, val: CellRichTextValue): void {
  const parts = val.richText.map((t) => ({ text: t.text.replace(/\n/g, '\r\n'), ts: t.font ? fontToStyle(t.font) : undefined }));
  const dataStream = parts.map((p) => p.text).join('') + '\r\n';
  const textRuns: SheetRichTextRun[] = [];
  let offset = 0;
  for (const part of parts) {
    if (part.ts) textRuns.push({ st: offset, ed: offset + part.text.length, ts: part.ts });
    offset += part.text.length;
  }
  uc.p = {
    id: RICH_TEXT_DOC_ID,
    body: { dataStream, ...(textRuns.length && { textRuns }) },
    documentStyle: {},
  };
}

/** exceljs run 字体 → CellStyle（复用单元格样式键名） */
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
```

- [ ] **Step 4.4: 跑测试确认通过**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 全部 PASS

- [ ] **Step 4.5: Commit**

```bash
git add src/core/xlsx/workbookToSnapshot.ts src/core/xlsx/xlsx.test.ts
git commit -m "feat: xlsx 导入富文本为 Univer cell.p 文档子集"
```

---

### Task 5: 导出方向 — 单元格样式

**Files:**
- Modify: `src/core/xlsx/snapshotToWorkbook.ts`
- Test: `src/core/xlsx/xlsx.test.ts`

- [ ] **Step 5.1: 写失败测试**

`xlsx.test.ts` 顶部 import 行改为（补充 CellStyle 类型）：

```ts
import type { CellStyle, WorkbookSnapshot } from '../../types/spreadsheet';
```

末尾追加 describe：

```ts
describe('xlsx 导出方向：扩展样式', () => {
  it('字体族/下划线/删除线/对齐/换行/边框写入 exceljs', () => {
    const snap = sampleSnapshot();
    snap.sheets.s1.cellData[0]![0]!.s = {
      ff: '微软雅黑', ul: { s: 1 }, st: { s: 1 },
      ht: 2, vt: 2, tb: 3,
      bd: {
        t: { s: 1, cl: { rgb: '#ff0000' } },
        b: { s: 7, cl: { rgb: '#00ff00' } },
        l: { s: 4, cl: { rgb: '#000000' } },
      },
    };
    const wb = new ExcelJS.Workbook();
    snapshotToWorkbook(snap, wb);
    const a1 = wb.worksheets[0].getCell('A1');
    expect(a1.font.name).toBe('微软雅黑');
    expect(a1.font.underline).toBe(true);
    expect(a1.font.strike).toBe(true);
    expect(a1.alignment).toMatchObject({ horizontal: 'center', vertical: 'middle', wrapText: true });
    expect(a1.border.top?.style).toBe('thin');
    expect(a1.border.top?.color?.argb).toBe('FFFF0000');
    expect(a1.border.bottom?.style).toBe('double');
    expect(a1.border.left?.style).toBe('dashed');
    expect(a1.border.left?.color?.argb).toBe('FF000000');
  });

  it('样式 id 形式的运行时快照同样走扩展样式导出', () => {
    const snap = sampleSnapshot();
    snap.styles = { sx: { ht: 3, vt: 3, tb: 3, ff: '宋体' } };
    snap.sheets.s1.cellData[1]![0]!.s = 'sx';
    const wb = new ExcelJS.Workbook();
    snapshotToWorkbook(snap, wb);
    const a2 = wb.worksheets[0].getCell('A2');
    expect(a2.alignment).toMatchObject({ horizontal: 'right', vertical: 'bottom', wrapText: true });
    expect(a2.font.name).toBe('宋体');
  });
});
```

- [ ] **Step 5.2: 跑测试确认失败**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 2 个新用例 FAIL（font.name undefined、alignment undefined）

- [ ] **Step 5.3: 实现 applyStyle 扩展**

`snapshotToWorkbook.ts` 顶部类型导入行改为：

```ts
import type { Cell, CellRichTextValue, CellValue, Workbook, Worksheet } from 'exceljs';
import type { CellStyle, SnapshotCell, SnapshotSheet, WorkbookSnapshot } from '../../types/spreadsheet';
```

常量区（`PX_TO_CHAR` 后）追加：

```ts
/** Univer BorderStyleTypes 数值枚举 → exceljs/OOXML 线型（1:1 直映） */
const BORDER_STYLE_NAMES: Record<number, string> = {
  1: 'thin', 2: 'hair', 3: 'dotted', 4: 'dashed', 5: 'dashDot', 6: 'dashDotDot', 7: 'double',
  8: 'medium', 9: 'mediumDashed', 10: 'mediumDashDot', 11: 'mediumDashDotDot',
  12: 'slantDashDot', 13: 'thick',
};
const H_ALIGN_NAMES: Record<number, string> = { 1: 'left', 2: 'center', 3: 'right' };
const V_ALIGN_NAMES: Record<number, string> = { 1: 'top', 2: 'middle', 3: 'bottom' };
const WRAP_STRATEGY_WRAP = 3;
```

`applyStyle` 整体替换为：

```ts
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
```

- [ ] **Step 5.4: 跑测试确认通过**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 全部 PASS

- [ ] **Step 5.5: Commit**

```bash
git add src/core/xlsx/snapshotToWorkbook.ts src/core/xlsx/xlsx.test.ts
git commit -m "feat: xlsx 导出支持边框/对齐/换行/字体族/下划线/删除线"
```

---

### Task 6: 导出方向 — 冻结窗格与隐藏行列

**Files:**
- Modify: `src/core/xlsx/snapshotToWorkbook.ts`
- Test: `src/core/xlsx/xlsx.test.ts`

- [ ] **Step 6.1: 写失败测试**

在 Task 5 的 describe 内追加：

```ts
  it('冻结窗格与隐藏行列写入 exceljs', () => {
    const snap = sampleSnapshot();
    snap.sheets.s1.freeze = { startRow: 2, startColumn: 1, xAxisSplit: 1, yAxisSplit: 2 };
    snap.sheets.s1.rowData = { 0: { h: 32 }, 3: { h: 0, hd: 1 } };
    snap.sheets.s1.columnData = { 0: { w: 120 }, 2: { w: 0, hd: 1 } };
    const wb = new ExcelJS.Workbook();
    snapshotToWorkbook(snap, wb);
    const ws = wb.worksheets[0];
    expect(ws.views?.[0]).toMatchObject({ state: 'frozen', xSplit: 1, ySplit: 2 });
    expect(ws.getRow(1).height).toBeCloseTo(24, 0);
    expect(ws.getRow(4).hidden).toBe(true);
    expect(ws.getColumn(1).width).toBeCloseTo(12, 0);
    expect(ws.getColumn(3).hidden).toBe(true);
  });
```

- [ ] **Step 6.2: 跑测试确认失败**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 1 个新用例 FAIL（views undefined、row4 未隐藏）

- [ ] **Step 6.3: 实现 writeSheet 扩展**

`writeSheet` 中三个写入循环段整体替换为：

```ts
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
```

文件末尾新增（`decodeRef` 的逆运算）：

```ts
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
```

- [ ] **Step 6.4: 跑测试确认通过**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 全部 PASS

- [ ] **Step 6.5: Commit**

```bash
git add src/core/xlsx/snapshotToWorkbook.ts src/core/xlsx/xlsx.test.ts
git commit -m "feat: xlsx 导出支持冻结窗格与隐藏行列"
```

---

### Task 7: 导出方向 — 富文本 cell.p → exceljs

**Files:**
- Modify: `src/core/xlsx/snapshotToWorkbook.ts`
- Test: `src/core/xlsx/xlsx.test.ts`

- [ ] **Step 7.1: 写失败测试**

在 Task 5 的 describe 内追加：

```ts
  it('cell.p 富文本写入 exceljs 富文本值', () => {
    const snap = sampleSnapshot();
    snap.sheets.s1.cellData[0]![0] = {
      p: {
        id: '__eflink-rich-text',
        documentStyle: {},
        body: {
          dataStream: '普通红色结尾\r\n',
          textRuns: [
            { st: 0, ed: 2, ts: { fs: 12 } },
            { st: 2, ed: 4, ts: { bl: 1, cl: { rgb: '#ff0000' } } },
          ],
        },
      },
    };
    const wb = new ExcelJS.Workbook();
    snapshotToWorkbook(snap, wb);
    const value = wb.worksheets[0].getCell('A1').value;
    expect(value).toMatchObject({
      richText: [
        { text: '普通', font: { size: 12 } },
        { text: '红色', font: { bold: true, color: { argb: 'FFFF0000' } } },
        { text: '结尾' },
      ],
    });
  });
```

- [ ] **Step 7.2: 跑测试确认失败**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 1 个新用例 FAIL（现状导出为 null）

- [ ] **Step 7.3: 实现 rich text 写出**

`toCellValue` 替换为：

```ts
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
```

注意：`\n`（多段落）在此方向还原为 exceljs 文本内换行，exceljs 写 xlsx 时保留在字符串里，再导入时由 `applyRichText` 重新切段——roundtrip 自洽。

- [ ] **Step 7.4: 跑测试确认通过**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 全部 PASS

- [ ] **Step 7.5: Commit**

```bash
git add src/core/xlsx/snapshotToWorkbook.ts src/core/xlsx/xlsx.test.ts
git commit -m "feat: xlsx 导出 Univer 富文本为 exceljs 富文本值"
```

---

### Task 8: 全链路 roundtrip 综合测试

**Files:**
- Test: `src/core/xlsx/xlsx.test.ts`

- [ ] **Step 8.1: 写 roundtrip 测试**

在 `xlsx.test.ts` 的 `describe('xlsx 双向转换')` 内追加：

```ts
  it('扩展样式/冻结/隐藏/富文本 全链路 roundtrip 无损', () => {
    const snap = sampleSnapshot();
    const s1 = snap.sheets.s1;
    s1.cellData[0]![1]!.s = {
      ff: '微软雅黑', ul: { s: 1 }, st: { s: 1 }, ht: 2, vt: 2, tb: 3,
      bd: {
        t: { s: 1, cl: { rgb: '#ff0000' } },
        b: { s: 7, cl: { rgb: '#00ff00' } },
        l: { s: 4, cl: { rgb: '#000000' } },
        r: { s: 13, cl: { rgb: '#0000ff' } },
      },
    };
    s1.cellData[6] = {
      0: {
        p: {
          id: '__eflink-rich-text',
          documentStyle: {},
          body: {
            dataStream: '第一段\r\n第二段\r\n',
            textRuns: [{ st: 3, ed: 6, ts: { bl: 1, cl: { rgb: '#ff0000' } } }],
          },
        },
      },
    };
    s1.freeze = { startRow: 2, startColumn: 1, xAxisSplit: 1, yAxisSplit: 2 };
    s1.rowData = { 0: { h: 32 }, 3: { h: 0, hd: 1 } };
    s1.columnData = { 0: { w: 120 }, 2: { w: 0, hd: 1 } };

    const out = roundTrip(snap);
    const os1 = Object.values(out.sheets)[0];
    const st = os1.cellData[0]?.[1]?.s as CellStyle;
    expect(st.ff).toBe('微软雅黑');
    expect(st.ul).toEqual({ s: 1 });
    expect(st.st).toEqual({ s: 1 });
    expect(st.ht).toBe(2);
    expect(st.vt).toBe(2);
    expect(st.tb).toBe(3);
    expect(st.bd?.t).toEqual({ s: 1, cl: { rgb: '#ff0000' } });
    expect(st.bd?.b).toEqual({ s: 7, cl: { rgb: '#00ff00' } });
    expect(st.bd?.l).toEqual({ s: 4, cl: { rgb: '#000000' } });
    expect(st.bd?.r).toEqual({ s: 13, cl: { rgb: '#0000ff' } });
    expect(os1.freeze).toEqual({ startRow: 2, startColumn: 1, xAxisSplit: 1, yAxisSplit: 2 });
    expect(os1.rowData?.[3]?.hd).toBe(1);
    expect(os1.columnData?.[2]?.hd).toBe(1);
    const p = os1.cellData[6]?.[0]?.p;
    expect(p?.body.dataStream.replace(/\r\n/g, '')).toBe('第一段第二段');
    expect(p?.body.textRuns?.[0]).toMatchObject({ st: 3, ed: 6, ts: { bl: 1, cl: { rgb: '#ff0000' } } });
  });
```

- [ ] **Step 8.2: 跑测试**

Run: `npx vitest run src/core/xlsx/xlsx.test.ts`
Expected: 全部 PASS（若富文本段间 run 区间有 ±1 漂移，修正 `richTextToCellValue`/`applyRichText` 后重跑至 PASS，不得改断言迁就 bug）

- [ ] **Step 8.3: Commit**

```bash
git add src/core/xlsx/xlsx.test.ts
git commit -m "test: xlsx 扩展样式/冻结/隐藏/富文本全链路 roundtrip"
```

---

### Task 9: editorStore.reloadToken + SheetEditor 重载

**Files:**
- Modify: `src/store/editorStore.ts`
- Modify: `src/components/SheetEditor.tsx:70-106`
- Test: `src/store/editorStore.test.ts`（新建）

- [ ] **Step 9.1: 写失败测试**

新建 `src/store/editorStore.test.ts`：

```ts
import { describe, expect, it } from 'vitest';
import { useEditorStore } from './editorStore';

describe('editorStore reloadToken', () => {
  it('bumpReload 递增重载计数', () => {
    const before = useEditorStore.getState().reloadToken;
    useEditorStore.getState().bumpReload();
    expect(useEditorStore.getState().reloadToken).toBe(before + 1);
  });
});
```

- [ ] **Step 9.2: 跑测试确认失败**

Run: `npx vitest run src/store/editorStore.test.ts`
Expected: FAIL（`reloadToken` 不存在）

- [ ] **Step 9.3: 实现 store 字段**

`editorStore.ts` 的 `EditorStore` 接口追加（`markSaved` 之后）：

```ts
  /** 导入替换等场景：递增触发 SheetEditor 重新加载当前文档（docId 不变内容已换） */
  reloadToken: number;
  bumpReload: () => void;
```

`create` 实现体追加（`markSaved` 之后）：

```ts
  reloadToken: 0,
  bumpReload: () => set((s) => ({ reloadToken: s.reloadToken + 1 })),
```

- [ ] **Step 9.4: SheetEditor 订阅重载**

`SheetEditor.tsx` 组件体内（`const storeDocId` 之类的既有订阅附近）新增订阅：

```tsx
  const reloadToken = useEditorStore((s) => s.reloadToken);
```

文档加载 effect（约 70 行处）的依赖数组由 `[docId, storage]` 改为 `[docId, storage, reloadToken]`。

- [ ] **Step 9.5: 跑测试确认通过**

Run: `npx vitest run src/store/editorStore.test.ts && npx vitest run`
Expected: 全部 PASS（SheetEditor 改动无专项测试，行为在 Task 12 demo 手动验收）

- [ ] **Step 9.6: Commit**

```bash
git add src/store/editorStore.ts src/store/editorStore.test.ts src/components/SheetEditor.tsx
git commit -m "feat: editorStore 增加 reloadToken 支持导入替换后编辑器重载"
```

---

### Task 10: filePicker + 导入导出动作（fileActions）

**Files:**
- Create: `src/core/filePicker.ts`
- Modify: `src/core/xlsx/fileIO.ts:48-50`（导出 trimExt）
- Modify: `src/core/fileActions.ts`
- Test: `src/core/filePicker.test.ts`（新建）
- Test: `src/core/fileActions.test.ts`

- [ ] **Step 10.1: 写失败测试（filePicker）**

新建 `src/core/filePicker.test.ts`：

```ts
// @vitest-environment jsdom
import { describe, expect, it, vi, afterEach } from 'vitest';
import { pickFile } from './filePicker';

afterEach(() => {
  vi.restoreAllMocks();
});

describe('filePicker 文件选择', () => {
  it('change 事件带文件时返回该文件，input 随后移除', async () => {
    const promise = pickFile('.xlsx,.xlsm');
    const input = document.querySelector<HTMLInputElement>('input[type="file"]')!;
    expect(input.accept).toBe('.xlsx,.xlsm');
    const file = new File(['x'], '报表.xlsx');
    Object.defineProperty(input, 'files', { value: [file] });
    input.dispatchEvent(new Event('change'));
    await expect(promise).resolves.toBe(file);
    expect(document.querySelector('input[type="file"]')).toBeNull();
  });

  it('cancel 事件返回 null', async () => {
    const promise = pickFile('.xlsx');
    const input = document.querySelector<HTMLInputElement>('input[type="file"]')!;
    input.dispatchEvent(new Event('cancel'));
    await expect(promise).resolves.toBeNull();
  });
});
```

- [ ] **Step 10.2: 跑测试确认失败**

Run: `npx vitest run src/core/filePicker.test.ts`
Expected: FAIL（模块不存在）

- [ ] **Step 10.3: 实现 filePicker**

新建 `src/core/filePicker.ts`：

```ts
/** 打开文件选择框（临时 input，不依赖 React ref）；用户取消返回 null */
export function pickFile(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    let settled = false;
    const finish = (file: File | null) => {
      if (settled) return;
      settled = true;
      input.remove();
      resolve(file);
    };
    input.addEventListener('change', () => finish(input.files?.[0] ?? null));
    input.addEventListener('cancel', () => finish(null));
    // 兜底：旧浏览器无 cancel 事件，选择框关闭（窗口重新聚焦）且无文件时视为取消
    window.addEventListener('focus', () => {
      window.setTimeout(() => finish(input.files?.[0] ?? null), 500);
    }, { once: true });
    document.body.appendChild(input);
    input.click();
  });
}
```

- [ ] **Step 10.4: 跑测试确认通过**

Run: `npx vitest run src/core/filePicker.test.ts`
Expected: 2 用例 PASS

- [ ] **Step 10.5: fileIO 导出 trimExt**

`src/core/xlsx/fileIO.ts` 末尾把 `function trimExt(...)` 加上 `export`：

```ts
export function trimExt(filename: string): string {
  return filename.replace(/\.(xlsx|xlsm|xls)$/i, '');
}
```

- [ ] **Step 10.6: 写失败测试（fileActions）**

`fileActions.test.ts` 的 `vi.mock('./xlsx/fileIO', ...)` 工厂替换为：

```ts
vi.mock('./xlsx/fileIO', () => ({
  exportPng: vi.fn(),
  exportXlsx: vi.fn(),
  importXlsx: vi.fn(),
  trimExt: (name: string) => name.replace(/\.(xlsx|xlsm|xls)$/i, ''),
}));

vi.mock('./filePicker', () => ({
  pickFile: vi.fn(),
}));
```

import 区补齐：

```ts
import { exportXlsx, importXlsx } from './xlsx/fileIO';
import { pickFile } from './filePicker';
import {
  exportEfexcelAction,
  exportPngAction,
  exportXlsxAction,
  importFileAction,
  importXlsxAction,
  importXlsxMenuAction,
  newDocAction,
  saveAction,
} from './fileActions';

const mockExportXlsx = vi.mocked(exportXlsx);
const mockImportXlsx = vi.mocked(importXlsx);
const mockPickFile = vi.mocked(pickFile);
```

describe 内追加用例：

```ts
  it('importXlsxAction 确认后替换当前文档内容并触发重载', async () => {
    const base = createDocument('基础文档');
    base.id = 'doc-base';
    await getDefaultStorage().save(base);
    mockImportXlsx.mockResolvedValueOnce(createEmptySnapshot('导入表'));

    const action = importXlsxAction(new File(['x'], '外部报表.xlsx'));
    expect(useUiStore.getState().confirm?.title).toBe('导入 Excel');
    expect(useUiStore.getState().confirm?.message).toContain('基础文档');
    useUiStore.getState().resolveConfirm(true);
    await action;

    expect(mockImportXlsx).toHaveBeenCalledTimes(1);
    const doc = await getDefaultStorage().load('doc-base');
    expect(doc?.title).toBe('外部报表');
    expect(doc?.snapshot.name).toBe('导入表');
    expect(useEditorStore.getState().reloadToken).toBe(1);
    expect(useUiStore.getState().toast).toBe('导入成功');
  });

  it('importXlsxAction 取消确认不解析不落盘', async () => {
    const action = importXlsxAction(new File(['x'], '外部报表.xlsx'));
    useUiStore.getState().resolveConfirm(false);
    await action;

    expect(mockImportXlsx).not.toHaveBeenCalled();
    expect(useEditorStore.getState().reloadToken).toBe(0);
  });

  it('importXlsxAction 解析失败提示导入失败且不触发重载', async () => {
    mockImportXlsx.mockRejectedValueOnce(new Error('解析失败'));
    const action = importXlsxAction(new File(['x'], '坏文件.xlsx'));
    useUiStore.getState().resolveConfirm(true);
    await action;

    expect(useUiStore.getState().toast).toBe('导入失败：无法解析该文件');
    expect(useEditorStore.getState().reloadToken).toBe(0);
  });

  it('exportXlsxAction 先保存再以当前文档快照触发导出', async () => {
    const doc = createDocument('基础文档');
    doc.id = 'doc-base';
    await getDefaultStorage().save(doc);
    mockCurrentSnapshot.mockReturnValueOnce(createEmptySnapshot());

    await exportXlsxAction();

    expect(mockSaveNow).toHaveBeenCalledTimes(1);
    expect(mockExportXlsx).toHaveBeenCalledTimes(1);
    expect(mockExportXlsx.mock.calls[0]?.[1]).toBe('基础文档');
    expect(useUiStore.getState().toast).toBe('已导出 xlsx');
  });

  it('importXlsxMenuAction 未选择文件时直接返回不弹确认', async () => {
    mockPickFile.mockResolvedValueOnce(null);
    await importXlsxMenuAction();
    expect(useUiStore.getState().confirm).toBeNull();
    expect(mockImportXlsx).not.toHaveBeenCalled();
  });

  it('importXlsxMenuAction 选中文件后走导入动作', async () => {
    const file = new File(['x'], '菜单报表.xlsx');
    mockPickFile.mockResolvedValueOnce(file);
    mockImportXlsx.mockResolvedValueOnce(createEmptySnapshot());

    const action = importXlsxMenuAction();
    useUiStore.getState().resolveConfirm(true);
    await action;

    expect(mockImportXlsx).toHaveBeenCalledWith(file);
  });
```

- [ ] **Step 10.7: 跑测试确认失败**

Run: `npx vitest run src/core/fileActions.test.ts`
Expected: 新用例 FAIL（importXlsxAction 等导出不存在）

- [ ] **Step 10.8: 实现 fileActions 动作**

`fileActions.ts` 顶部 import 改为：

```ts
import { currentSnapshot, saveNow } from './saveService';
import { exportEfexcel, importEfexcel } from './efexcel';
import { exportPng, exportXlsx, importXlsx, trimExt } from './xlsx/fileIO';
import { pickFile } from './filePicker';
import { getDefaultStorage } from '../storage/registry';
import { clearDraft } from '../storage/draft';
import { useDocumentsStore } from '../store/documentsStore';
import { useEditorStore } from '../store/editorStore';
import { useUiStore } from '../store/uiStore';
```

文件末尾追加：

```ts
/** 导出 Excel (.xlsx)：先保存，再取最新快照转换下载 */
export async function exportXlsxAction(): Promise<void> {
  const { docId } = useEditorStore.getState();
  if (!docId || !currentSnapshot()) return;
  try {
    await saveNow();
    const doc = await getDefaultStorage().load(docId);
    if (!doc) throw new Error(`文档不存在: ${docId}`);
    await exportXlsx(doc.snapshot, doc.title);
    useUiStore.getState().showToast('已导出 xlsx');
  } catch (err) {
    console.error('导出 xlsx 失败', err);
    useUiStore.getState().showToast('导出失败');
  }
}

/** 导入 Excel (.xlsx)：解析成功后替换当前文档内容（docId 不变），并触发编辑器重载 */
export async function importXlsxAction(file: File): Promise<void> {
  const { docId, title } = useEditorStore.getState();
  if (!docId) return;
  const ok = await useUiStore.getState().requestConfirm({
    title: '导入 Excel',
    message: `导入将替换当前文档「${title}」的内容，确定继续吗？`,
  });
  if (!ok) return;
  try {
    // 先解析：失败在落盘前抛出，当前文档不受影响
    const snapshot = await importXlsx(file);
    await getDefaultStorage().updateContent(docId, trimExt(file.name), snapshot);
    clearDraft(docId); // 防重载读到旧草稿
    void useDocumentsStore.getState().refresh();
    useEditorStore.getState().bumpReload();
    useUiStore.getState().showToast('导入成功');
  } catch (err) {
    console.error('导入 Excel 失败', err);
    useUiStore.getState().showToast('导入失败：无法解析该文件');
  }
}

/** 菜单入口：弹文件选择框（限 xlsx/xlsm），选中后走导入动作 */
export async function importXlsxMenuAction(): Promise<void> {
  const file = await pickFile('.xlsx,.xlsm');
  if (!file) return;
  await importXlsxAction(file);
}
```

- [ ] **Step 10.9: 跑测试确认通过**

Run: `npx vitest run src/core/fileActions.test.ts`
Expected: 全部 PASS

- [ ] **Step 10.10: Commit**

```bash
git add src/core/filePicker.ts src/core/filePicker.test.ts src/core/xlsx/fileIO.ts src/core/fileActions.ts src/core/fileActions.test.ts
git commit -m "feat: xlsx 导入(替换当前文档)/导出文件动作与菜单文件选择器"
```

---

### Task 11: MenuDropdown 菜单项挂载

**Files:**
- Modify: `src/components/toolbar/MenuDropdown.tsx:7-42`
- Modify: `src/components/toolbar/AppMenu.test.tsx:14-20`

- [ ] **Step 11.1: 更新 AppMenu 测试 mock（先行，保证 render 不炸）**

`AppMenu.test.tsx` 的 `vi.mock('../../core/fileActions', ...)` 工厂替换为：

```ts
vi.mock('../../core/fileActions', () => ({
  exportPngAction: vi.fn(),
  exportEfexcelAction: vi.fn(),
  exportXlsxAction: vi.fn(),
  importFileAction: vi.fn(),
  importXlsxMenuAction: vi.fn(),
  newDocAction: vi.fn(),
  saveAction: vi.fn(),
}));
```

- [ ] **Step 11.2: MenuDropdown 新增菜单项**

`MenuDropdown.tsx` 的 fileActions import 改为：

```ts
import {
  exportEfexcelAction,
  exportPngAction,
  exportXlsxAction,
  importXlsxMenuAction,
  newDocAction,
  saveAction,
} from '../../core/fileActions';
```

`useMenuItems` 的 `fileChildren` 替换为：

```ts
  const fileChildren: MenuLeaf[] = [
    { label: '新建表格', iconClass: 'myf-icon-add', action: () => void newDocAction() },
    { label: '保存', icon: <Save size={16} />, shortcut: '⌘S', action: () => void saveAction() },
    { label: '导入表格(.efexcel)', iconClass: 'myf-icon-file-import', action: () => importRef.current?.click() },
    { label: '导入 Excel(.xlsx)', iconClass: 'myf-icon-file-import', action: () => void importXlsxMenuAction() },
    { label: '导出表格(.efexcel)', iconClass: 'myf-icon-file-export', action: () => void exportEfexcelAction() },
    { label: '导出 Excel(.xlsx)', iconClass: 'myf-icon-file-export', action: () => void exportXlsxAction() },
    { label: '导出图片', iconClass: 'myf-icon-file-download', action: () => void exportPngAction() },
  ];
```

- [ ] **Step 11.3: 跑组件测试确认通过**

Run: `npx vitest run src/components/toolbar/AppMenu.test.tsx`
Expected: 全部 PASS（既有用例不受影响）

- [ ] **Step 11.4: Commit**

```bash
git add src/components/toolbar/MenuDropdown.tsx src/components/toolbar/AppMenu.test.tsx
git commit -m "feat: 文件菜单挂载 xlsx 导入/导出入口"
```

---

### Task 12: 全量验证 + 手动验收

**Files:** 无新增（验证任务）

- [ ] **Step 12.1: 全量测试**

Run: `npx vitest run`
Expected: 全部 PASS，无失败/跳过异常

- [ ] **Step 12.2: Lint 与类型检查**

Run: `npm run lint && npx tsc -b --noEmit`
Expected: 无错误（lint 无新增告警）

- [ ] **Step 12.3: 构建**

Run: `npm run build`
Expected: 构建成功

- [ ] **Step 12.4: demo 手动验收**

Run: `cd /Users/apple/Documents/myf-project/eflink.tech/eflink-excel && npm run dev:demo`，浏览器打开 demo 地址，逐项验证：

1. 用 Excel/WPS 制作含以下特征的真实 xlsx：标题加粗+背景色+边框、居中表头、合并单元格、冻结首行、一个隐藏列、一个富文本单元格、一个公式（如 SUM）
2. 「文件 → 导入 Excel(.xlsx)」→ 确认弹窗 → 导入后网格内容/边框/合并/冻结/隐藏列/富文本/公式显示正确（公式重算值正确，Excel 独有函数允许 #NAME?）
3. 编辑一格 → 「文件 → 导出 Excel(.xlsx)」→ 用 Excel/WPS 打开导出文件核对上述特征
4. 导入时取消确认弹窗 → 当前文档内容不变；选一个非 xlsx 文件 → toast「导入失败：无法解析该文件」且页面正常

- [ ] **Step 12.5: 提交文档与收尾**

```bash
cd /Users/apple/Documents/myf-project/eflink.tech/eflink-excel
git add docs/superpowers/specs/2026-09-17-xlsx-import-export-design.md docs/superpowers/plans/2026-09-17-xlsx-import-export.md
git commit -m "docs: xlsx 导入导出设计修正（边框全量直映/数值枚举）与实施计划"
```

---

## 已知边界（实现者须知）

- 公式：保留字符串由 Univer 重算；Excel 独有函数显示 `#NAME?` 是**已接受的既定策略**，不做逐格回退
- 富文本只保段落文字 + run 字体样式；段内图片/上下标/链接不支持；Univer 对 `cell.p` 的渲染要求以 demo 实测为准（Task 12.4 是它的验收关）
- 列宽沿用字符宽↔px 经验换算，不做字体度量级精确
- `.xls` 老格式不支持（exceljs 限制），菜单选择器已限 `.xlsx,.xlsm`
- 超链接放二期
