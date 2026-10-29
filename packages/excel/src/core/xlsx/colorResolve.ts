// OOXML 颜色解析：把 exceljs 暴露的 {argb} / {theme,tint} / {indexed} 三种颜色引用
// 统一解析成 Univer 需要的 '#rrggbb'。真实 Excel/WPS/LibreOffice 文件里，字体色、填充色、
// 边框色绝大多数走「主题色 + tint」或「索引色」，而非硬编码 argb——只读 argb 会让这些颜色全丢。
// 本模块保持纯逻辑（不依赖 DOM/exceljs 运行时），供 node 环境单测。
import { findFirst, parseXml, type XNode } from './xmlLite';

/** exceljs 的 Color 类型只声明了 argb/theme，运行时还带 indexed/tint（类型缺陷），这里按运行时形状读 */
export interface ColorRef {
  argb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

/** OOXML 传统索引色板（ECMA-376 §18.8.27）；64=窗口文本(auto→黑)，65=窗口背景(→白) */
export const INDEXED_PALETTE: string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
  '000000', 'FFFFFF',
];

/** clrScheme 在 theme1.xml 里的文档顺序（dk1 在前、lt1 在后……） */
const CLR_SCHEME_SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
/** 主题索引 → clrScheme 文档下标：0/1 与 2/3 两对是「浅/深」互换（0=lt1、1=dk1、2=lt2、3=dk2），accent 起顺延 */
const THEME_INDEX_TO_SLOT = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11];

/** 缺省 Office 主题（按主题索引序），theme1.xml 缺失或解析失败时兜底 */
export const STANDARD_OFFICE_THEME: string[] = [
  'FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646', '0000FF', '800080',
];

/** 归一化成 6 位大写 RGB：去 #、8 位 ARGB 去 alpha、3 位展开、异常回退黑 */
function normalizeHex(hex: string): string {
  let v = hex.replace('#', '').toUpperCase();
  if (v.length === 8) v = v.slice(2);
  if (v.length === 3) v = [...v].map((c) => c + c).join('');
  return /^[0-9A-F]{6}$/.test(v) ? v : '000000';
}

/**
 * 应用 OOXML tint（§20.1.2.3.14）：在 HLS 亮度空间线性调整 L 后转回 RGB。
 * tint>0 提亮（L'=L*(1-tint)+tint），tint<0 压暗（L'=L*(1+tint)）；用亮度空间才与 Excel 一致
 * （逐通道近似在压暗时会明显偏色，如 accent1「深 25%」Excel 为 #376092，逐通道会得 #3B618E）。
 */
export function applyTint(hex: string, tint?: number): string {
  const h = normalizeHex(hex);
  if (!tint) return h;
  const t = Math.max(-1, Math.min(1, tint));
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l0 = (max + min) / 2;
  const d = max - min;
  let hue = 0;
  let sat = 0;
  if (d !== 0) {
    sat = l0 < 0.5 ? d / (max + min) : d / (2 - max - min);
    if (max === r) hue = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) hue = (b - r) / d + 2;
    else hue = (r - g) / d + 4;
    hue /= 6;
  }
  const l = t < 0 ? l0 * (1 + t) : l0 * (1 - t) + t;
  const to255 = (x: number) => Math.max(0, Math.min(255, Math.round(x * 255)));
  const chan = (v: number) => v.toString(16).padStart(2, '0');
  if (sat === 0) {
    const gray = chan(to255(l));
    return (gray + gray + gray).toUpperCase();
  }
  const q = l < 0.5 ? l * (1 + sat) : l + sat - l * sat;
  const p = 2 * l - q;
  const hueToRgb = (tt: number) => {
    let x = tt;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  return (chan(to255(hueToRgb(hue + 1 / 3))) + chan(to255(hueToRgb(hue))) + chan(to255(hueToRgb(hue - 1 / 3)))).toUpperCase();
}

/** 单个 clrScheme 槽（如 <a:dk1>）→ 6 位大写 RGB；srgbClr 取 val，sysClr 取 lastClr（回退 val） */
function slotHex(slot: XNode | null): string | undefined {
  if (!slot) return undefined;
  const srgb = findFirst(slot, 'srgbClr');
  if (srgb?.attrs.val) return normalizeHex(srgb.attrs.val);
  const sys = findFirst(slot, 'sysClr');
  if (sys?.attrs.lastClr) return normalizeHex(sys.attrs.lastClr);
  if (sys?.attrs.val && /^[0-9A-Fa-f]{6}$/.test(sys.attrs.val)) return normalizeHex(sys.attrs.val);
  return undefined;
}

/** theme1.xml → 12 色主题板（按主题索引序）；结构不认识时返回 undefined 交调用方兜底 */
export function parseThemePalette(xml: string): string[] | undefined {
  if (!xml) return undefined;
  const scheme = findFirst(parseXml(xml), 'clrScheme');
  if (!scheme) return undefined;
  const bySlot = new Map<string, string>();
  for (const child of scheme.children) {
    const hex = slotHex(child);
    if (hex) bySlot.set(child.name, hex);
  }
  if (!bySlot.size) return undefined;
  return THEME_INDEX_TO_SLOT.map((slotIdx, themeIdx) => bySlot.get(CLR_SCHEME_SLOTS[slotIdx]) ?? STANDARD_OFFICE_THEME[themeIdx]);
}

/**
 * 颜色引用 → '#rrggbb'（小写，与既有 fromArgb 输出一致）。
 * argb 直接用；theme 按主题板取基色再套 tint；indexed 查索引色板；都无则 undefined（调用方决定是否补默认色）。
 */
export function resolveColorHex(ref: ColorRef | undefined, palette: string[]): string | undefined {
  if (!ref) return undefined;
  if (ref.argb) return '#' + normalizeHex(ref.argb).toLowerCase();
  if (typeof ref.theme === 'number') {
    const base = palette[ref.theme] ?? STANDARD_OFFICE_THEME[ref.theme];
    return base ? '#' + applyTint(base, ref.tint).toLowerCase() : undefined;
  }
  if (typeof ref.indexed === 'number') {
    const base = INDEXED_PALETTE[ref.indexed];
    return base ? '#' + base.toLowerCase() : undefined;
  }
  return undefined;
}
