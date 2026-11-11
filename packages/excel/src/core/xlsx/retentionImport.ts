// xlsx 条件格式 / 数据验证导入（解析侧）：exceljs 不保留这两类能力，直接在 zip 内读原始 XML。
// 与批注/图表一样，解析结果只产出纯数据，Univer 挂载后经 facade/命令落位（见 univer/importedRetentions.ts）。
// 本模块保持纯逻辑（不依赖 DOM/Univer 运行时），供 node 环境单测。
import JSZip from 'jszip';
import type { IRange } from '@univerjs/core';
import type { IConditionalFormattingRuleConfig } from '@univerjs/sheets-conditional-formatting';
import { findAll, findFirst, parseXml, type XNode } from './xmlLite';
import { listSheets } from './zipNav';
import { parseThemePalette, resolveColorHex, type ColorRef } from './colorResolve';

/** 条件格式样式：与 Univer IStyleBase 对齐（只保留 Excel dxf 能表达的字段） */
export interface ImportedCfStyle {
  bl?: 1;
  it?: 1;
  ul?: { s: 1 };
  st?: { s: 1 };
  cl?: { rgb: string };
  bg?: { rgb: string };
}

/**
 * 图标集先在解析侧保留 Excel 的图标名称与阈值，落位时再由 facade 侧换算成
 * Univer 的 iconId（iconMap 属于运行时资源，不适合放在纯解析模块）。
 */
export interface ImportedIconSetDescriptor {
  type: 'iconSet';
  isShowValue: boolean;
  iconType: string;
  values: { type: string; value?: number }[];
}

/** 解析侧产出：普通规则直接是 Univer 配置，图标集用上面的描述结构延后换算 */
export type ImportedCfRuleConfig = IConditionalFormattingRuleConfig | ImportedIconSetDescriptor;

export interface ImportedCfRule {
  sheetName: string;
  ranges: IRange[];
  /** Excel priority：数值越小优先级越高，落位时按升序应用 */
  priority: number;
  stopIfTrue: boolean;
  rule: ImportedCfRuleConfig;
}

export interface ImportedDvRule {
  sheetName: string;
  ranges: IRange[];
  type: string;
  operator?: string;
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  /** Univer 语义：true 显示下拉箭头；Excel 的 showDropDown 属性含义相反，解析时已翻转 */
  showDropDown?: boolean;
  showErrorMessage?: boolean;
  errorStyle?: number;
  errorTitle?: string;
  error?: string;
  showInputMessage?: boolean;
  promptTitle?: string;
  prompt?: string;
}

export interface ImportedRetentions {
  conditionalFormats: ImportedCfRule[];
  dataValidations: ImportedDvRule[];
}

// ---------- A1 区域 ----------

const CELL_REF = /^\$?([A-Za-z]+)\$?(\d+)$/;
const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

function columnNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}

/** "A1" / "A1:C9" / "$A$1:$C$9" → 0 基矩形；无法解析返回 null */
function parseArea(area: string): IRange | null {
  const parts = area.trim().split(':');
  const start = CELL_REF.exec(parts[0]);
  if (!start) return null;
  const row1 = Number(start[2]) - 1;
  const col1 = columnNumber(start[1]) - 1;
  if (row1 < 0 || col1 < 0) return null;
  const end = parts.length === 1 ? start : CELL_REF.exec(parts[1] ?? '');
  if (!end) return null;
  const row2 = Number(end[2]) - 1;
  const col2 = columnNumber(end[1]) - 1;
  if (row2 < 0 || col2 < 0) return null;
  return {
    startRow: Math.min(row1, row2),
    startColumn: Math.min(col1, col2),
    endRow: Math.min(Math.max(row1, row2), MAX_ROW - 1),
    endColumn: Math.min(Math.max(col1, col2), MAX_COL - 1),
  };
}

/** sqref 可含多个以空格分隔的区域（如 "A1:A10 C1:C10"） */
function parseSqref(sqref: string | undefined): IRange[] {
  if (!sqref) return [];
  return sqref
    .split(/\s+/)
    .map(parseArea)
    .filter((r): r is IRange => r !== null);
}

// ---------- dxf 样式 ----------

function colorOf(node: XNode | null, palette: string[]): { rgb: string } | undefined {
  if (!node) return undefined;
  const ref: ColorRef = {
    argb: node.attrs.rgb,
    theme: node.attrs.theme != null ? Number(node.attrs.theme) : undefined,
    tint: node.attrs.tint != null ? Number(node.attrs.tint) : undefined,
    indexed: node.attrs.indexed != null ? Number(node.attrs.indexed) : undefined,
  };
  const hex = resolveColorHex(ref, palette);
  return hex ? { rgb: hex } : undefined;
}

function flag(node: XNode | null): boolean {
  if (!node) return false;
  return node.attrs.val !== '0' && node.attrs.val !== 'false';
}

function parseDxf(dxf: XNode | undefined, palette: string[]): ImportedCfStyle {
  const style: ImportedCfStyle = {};
  if (!dxf) return style;
  const font = findFirst(dxf, 'font');
  if (font) {
    if (flag(findFirst(font, 'b'))) style.bl = 1;
    if (flag(findFirst(font, 'i'))) style.it = 1;
    if (flag(findFirst(font, 'strike'))) style.st = { s: 1 };
    if (flag(findFirst(font, 'u'))) style.ul = { s: 1 };
    const cl = colorOf(findFirst(font, 'color'), palette);
    if (cl) style.cl = cl;
  }
  const fill = findFirst(dxf, 'fill');
  if (fill) {
    const bg = colorOf(findFirst(fill, 'bgColor'), palette) ?? colorOf(findFirst(fill, 'fgColor'), palette);
    if (bg) style.bg = bg;
  }
  return style;
}

// ---------- 条件格式规则 ----------

interface CfValue {
  type: string;
  value?: number;
}

function cfValues(node: XNode | null): CfValue[] {
  if (!node) return [];
  return findAll(node, 'cfvo').map((el) => {
    const type = el.attrs.type ?? 'num';
    const raw = el.attrs.val;
    const value = raw == null ? undefined : Number(raw);
    return Number.isFinite(value) ? { type, value } : { type };
  });
}

function cfColors(node: XNode | null, palette: string[]): string[] {
  if (!node) return [];
  return findAll(node, 'color')
    .map((el) => colorOf(el, palette)?.rgb)
    .filter((c): c is string => Boolean(c));
}

function numeric(formula: string | undefined): number | null {
  if (formula == null) return null;
  const trimmed = formula.trim();
  const n = Number(trimmed);
  return trimmed !== '' && Number.isFinite(n) ? n : null;
}

function unquote(value: string): string {
  const trimmed = value.trim();
  return trimmed.length >= 2 && trimmed.startsWith('"') && trimmed.endsWith('"')
    ? trimmed.slice(1, -1).replace(/""/g, '"')
    : trimmed;
}

const NUMBER_OPERATORS = new Set([
  'between', 'notBetween', 'equal', 'notEqual',
  'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual',
]);

const TEXT_OPERATORS = new Set([
  'beginsWith', 'endsWith', 'containsText', 'notContainsText', 'equal', 'notEqual',
  'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors',
]);

const TIME_PERIODS = new Set([
  'today', 'yesterday', 'tomorrow', 'last7Days',
  'thisMonth', 'lastMonth', 'nextMonth',
  'thisWeek', 'lastWeek', 'nextWeek',
]);

function numberRule(operator: string, formulas: string[], style: ImportedCfStyle): ImportedCfRuleConfig | null {
  if (!NUMBER_OPERATORS.has(operator)) return null;
  const a = numeric(formulas[0]);
  const b = numeric(formulas[1]);
  if (a == null) {
    // 非数值比较退化为文本相等/不等；其余算子无从表达，跳过
    if (operator === 'equal' || operator === 'notEqual') {
      return { type: 'highlightCell', subType: 'text', operator, value: unquote(formulas[0] ?? ''), style } as IConditionalFormattingRuleConfig;
    }
    return null;
  }
  if (operator === 'between' || operator === 'notBetween') {
    if (b == null) return null;
    return {
      type: 'highlightCell',
      subType: 'number',
      operator,
      value: [Math.min(a, b), Math.max(a, b)],
      style,
    } as IConditionalFormattingRuleConfig;
  }
  return { type: 'highlightCell', subType: 'number', operator, value: a, style } as IConditionalFormattingRuleConfig;
}

function highlightRule(cf: XNode, style: ImportedCfStyle, formulas: string[], palette: string[]): ImportedCfRuleConfig | null {
  const type = cf.attrs.type ?? '';
  const operator = cf.attrs.operator ?? '';
  switch (type) {
    case 'expression': {
      const value = formulas[0];
      if (!value) return null;
      return { type: 'highlightCell', subType: 'formula', value, style } as IConditionalFormattingRuleConfig;
    }
    case 'cellIs':
      return numberRule(operator, formulas, style);
    case 'containsText':
      if (!TEXT_OPERATORS.has(operator)) return null;
      return { type: 'highlightCell', subType: 'text', operator, value: cf.attrs.text ?? '', style } as IConditionalFormattingRuleConfig;
    case 'beginsWith':
    case 'endsWith':
      return { type: 'highlightCell', subType: 'text', operator: type, value: cf.attrs.text ?? '', style } as IConditionalFormattingRuleConfig;
    case 'notContainsText':
      return { type: 'highlightCell', subType: 'text', operator: 'notContainsText', value: cf.attrs.text ?? '', style } as IConditionalFormattingRuleConfig;
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      return { type: 'highlightCell', subType: 'text', operator: type, style } as IConditionalFormattingRuleConfig;
    case 'duplicateValues':
      return { type: 'highlightCell', subType: 'duplicateValues', style } as IConditionalFormattingRuleConfig;
    case 'uniqueValues':
      return { type: 'highlightCell', subType: 'uniqueValues', style } as IConditionalFormattingRuleConfig;
    case 'timePeriod': {
      const period = cf.attrs.timePeriod ?? '';
      if (!TIME_PERIODS.has(period)) return null;
      return { type: 'highlightCell', subType: 'timePeriod', operator: period, style } as IConditionalFormattingRuleConfig;
    }
    case 'top10': {
      const rank = numeric(cf.attrs.rank) ?? 10;
      return {
        type: 'highlightCell',
        subType: 'rank',
        isBottom: cf.attrs.bottom === '1',
        isPercent: cf.attrs.percent === '1',
        value: rank,
        style,
      } as IConditionalFormattingRuleConfig;
    }
    case 'aboveAverage': {
      // aboveAverage="0" 表示「低于平均值」；等于平均值的特例用 lessThanOrEqual
      const op = cf.attrs.aboveAverage === '0' ? 'lessThan' : 'greaterThan';
      return { type: 'highlightCell', subType: 'average', operator: op, style } as IConditionalFormattingRuleConfig;
    }
    case 'colorScale': {
      const scale = findFirst(cf, 'colorScale');
      const values = cfValues(scale);
      const colors = cfColors(scale, palette);
      const config = values
        .map((value, index) => ({ index, value, color: colors[index] ?? '#ffffff' }))
        .filter((entry) => entry.color);
      if (config.length < 2) return null;
      return { type: 'colorScale', config } as IConditionalFormattingRuleConfig;
    }
    case 'dataBar': {
      const bar = findFirst(cf, 'dataBar');
      const values = cfValues(bar);
      const color = cfColors(bar, palette)[0] ?? '#638ec6';
      return {
        type: 'dataBar',
        isShowValue: cf.attrs.showValue !== '0',
        config: {
          min: values[0] ?? { type: 'min' },
          max: values[1] ?? { type: 'max' },
          isGradient: bar?.attrs.gradient !== '0',
          positiveColor: color,
          nativeColor: color,
        },
      } as IConditionalFormattingRuleConfig;
    }
    case 'iconSet': {
      const set = findFirst(cf, 'iconSet');
      const values = cfValues(set);
      const iconType = set?.attrs.iconSet ?? '3TrafficLights1';
      if (!values.length) return null;
      return {
        type: 'iconSet',
        isShowValue: set?.attrs.showValue !== '0',
        iconType,
        values,
      };
    }
    default:
      return null;
  }
}

function readConditionalFormats(sheetXml: XNode | null, sheetName: string, dxfs: XNode[], palette: string[]): ImportedCfRule[] {
  const out: ImportedCfRule[] = [];
  for (const block of findAll(sheetXml, 'conditionalFormatting')) {
    const ranges = parseSqref(block.attrs.sqref);
    if (!ranges.length) continue;
    for (const cf of findAll(block, 'cfRule')) {
      const priority = numeric(cf.attrs.priority) ?? out.length + 1;
      const stopIfTrue = cf.attrs.stopIfTrue === '1';
      const formulas = findAll(cf, 'formula').map((f) => f.text.trim());
      const dxf = dxfs[Number(cf.attrs.dxfId ?? -1)];
      const style = parseDxf(dxf, palette);
      const rule = highlightRule(cf, style, formulas, palette);
      if (!rule) continue;
      out.push({ sheetName, ranges, priority, stopIfTrue, rule });
    }
  }
  return out.sort((a, b) => a.priority - b.priority);
}

// ---------- 数据验证 ----------

const DV_TYPES: Record<string, string> = {
  whole: 'whole',
  decimal: 'decimal',
  list: 'list',
  date: 'date',
  time: 'time',
  textLength: 'textLength',
  custom: 'custom',
};

const DV_OPERATORS = new Set([
  'between', 'notBetween', 'equal', 'notEqual',
  'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual',
]);

const DV_ERROR_STYLES: Record<string, number> = {
  information: 0,
  stop: 1,
  warning: 2,
};

function readDataValidations(sheetXml: XNode | null, sheetName: string): ImportedDvRule[] {
  const out: ImportedDvRule[] = [];
  for (const dv of findAll(sheetXml, 'dataValidation')) {
    const ranges = parseSqref(dv.attrs.sqref);
    if (!ranges.length) continue;
    const excelType = dv.attrs.type ?? '';
    const type = DV_TYPES[excelType];
    if (!type) continue; // none / 未知类型不落位
    const operator = dv.attrs.operator && DV_OPERATORS.has(dv.attrs.operator) ? dv.attrs.operator : undefined;
    const rawShowDropDown = dv.attrs.showDropDown;
    out.push({
      sheetName,
      ranges,
      type,
      ...(operator ? { operator } : {}),
      ...(findFirst(dv, 'formula1')?.text ? { formula1: findFirst(dv, 'formula1')!.text.trim() } : {}),
      ...(findFirst(dv, 'formula2')?.text ? { formula2: findFirst(dv, 'formula2')!.text.trim() } : {}),
      allowBlank: dv.attrs.allowBlank === '1',
      // Excel 的 showDropDown=1 表示隐藏下拉箭头，与 Univer 语义相反
      ...(rawShowDropDown != null ? { showDropDown: rawShowDropDown !== '1' } : {}),
      showErrorMessage: dv.attrs.showErrorMessage === '1',
      ...(dv.attrs.errorStyle ? { errorStyle: DV_ERROR_STYLES[dv.attrs.errorStyle] ?? 1 } : {}),
      ...(dv.attrs.errorTitle ? { errorTitle: dv.attrs.errorTitle } : {}),
      ...(dv.attrs.error ? { error: dv.attrs.error } : {}),
      showInputMessage: dv.attrs.showInputMessage === '1',
      ...(dv.attrs.promptTitle ? { promptTitle: dv.attrs.promptTitle } : {}),
      ...(dv.attrs.prompt ? { prompt: dv.attrs.prompt } : {}),
    });
  }
  return out;
}

// ---------- 对外入口 ----------

async function readDxfs(zip: JSZip): Promise<XNode[]> {
  const xml = await zip.file('xl/styles.xml')?.async('string');
  const root = xml ? parseXml(xml) : null;
  const container = findFirst(root, 'dxfs');
  if (!container) return [];
  return findFirst(container, 'dxf') ? findAll(container, 'dxf') : [];
}

export async function extractRetentions(buffer: ArrayBuffer): Promise<ImportedRetentions> {
  const zip = await JSZip.loadAsync(buffer);
  const themeXml = await zip.file('xl/theme/theme1.xml')?.async('string');
  const palette = (themeXml ? parseThemePalette(themeXml) : undefined) ?? [];
  const dxfs = await readDxfs(zip);
  const conditionalFormats: ImportedCfRule[] = [];
  const dataValidations: ImportedDvRule[] = [];

  for (const { sheetName, sheetPath } of await listSheets(zip)) {
    const xml = await zip.file(sheetPath)?.async('string');
    if (!xml) continue;
    const root = parseXml(xml);
    if (!root) continue;
    conditionalFormats.push(...readConditionalFormats(root, sheetName, dxfs, palette));
    dataValidations.push(...readDataValidations(root, sheetName));
  }
  return { conditionalFormats, dataValidations };
}

// ---------- 待落位队列 ----------

export type RetentionTask = ({ kind: 'cf' } & ImportedCfRule) | ({ kind: 'dv' } & ImportedDvRule);

let pendingRetentions: RetentionTask[] = [];

export function queueImportedRetentions(retentions: ImportedRetentions): void {
  pendingRetentions = [
    ...retentions.conditionalFormats.map((r) => ({ kind: 'cf' as const, ...r })),
    ...retentions.dataValidations.map((r) => ({ kind: 'dv' as const, ...r })),
  ];
}

export function peekImportedRetention(): RetentionTask | null {
  return pendingRetentions[0] ?? null;
}

export function shiftImportedRetention(): RetentionTask | null {
  return pendingRetentions.shift() ?? null;
}

export function unshiftImportedRetention(task: RetentionTask): void {
  pendingRetentions.unshift(task);
}
