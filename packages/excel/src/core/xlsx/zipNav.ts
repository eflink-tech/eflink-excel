// xlsx（zip 包）导航工具：图表（chartImport）与批注/超链接（annotationImport）解析共用。
// OOXML 的部件之间靠 *_rels/*.rels 关系文件互相引用，路径可能是相对的（含 ..），
// 这里统一做路径归一化与关系读取，避免各解析器各写一套。
import JSZip from 'jszip';
import { findAll, parseXml } from './xmlLite';

/** zip 内路径的目录部分（无 / 时为空串） */
export function dirname(path: string): string {
  const i = path.lastIndexOf('/');
  return i === -1 ? '' : path.slice(0, i);
}

/** rels 里的相对路径 → zip 内绝对路径（以 / 开头的包绝对路径也归一化） */
export function resolvePath(fromDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (fromDir ? fromDir.split('/') : []).concat(target.split('/'));
  const out: string[] = [];
  for (const p of parts) {
    if (!p || p === '.') continue;
    if (p === '..') out.pop();
    else out.push(p);
  }
  return out.join('/');
}

/** 部件路径 → 其 rels 路径（xl/worksheets/sheet1.xml → xl/worksheets/_rels/sheet1.xml.rels） */
export function relsPathOf(partPath: string): string {
  return `${dirname(partPath)}/_rels/${partPath.split('/').pop()}.rels`;
}

/** 解析 *_rels/*.rels 内容为 Id → Target 映射（Relationship 属性名首字母大写） */
export async function readRels(zip: JSZip, relsPath: string): Promise<Record<string, string>> {
  const file = zip.file(relsPath);
  if (!file) return {};
  const root = parseXml(await file.async('string'));
  const map: Record<string, string> = {};
  for (const rel of findAll(root, 'Relationship')) {
    map[rel.attrs.Id ?? rel.attrs.id ?? ''] = rel.attrs.Target ?? rel.attrs.target ?? '';
  }
  return map;
}

export interface SheetEntry {
  /** 工作表名（sheet 元素 name 属性，原样、未转义） */
  sheetName: string;
  /** 工作表 XML 在 zip 内的绝对路径 */
  sheetPath: string;
  /** 该工作表的 rels 路径（可能不存在，读取时按空处理） */
  sheetRelsPath: string;
}

/**
 * workbook.xml → 每个 sheet 的名称与 zip 内路径。
 * 关系缺失或指向的部件不在包内的 sheet 直接跳过（防御损坏/非常规文件）。
 */
export async function listSheets(zip: JSZip): Promise<SheetEntry[]> {
  const workbookXml = parseXml((await zip.file('xl/workbook.xml')?.async('string')) ?? '');
  const workbookRels = await readRels(zip, 'xl/_rels/workbook.xml.rels');
  const out: SheetEntry[] = [];
  for (const sheetEl of workbookXml ? findAll(workbookXml, 'sheet') : []) {
    const sheetName = sheetEl.attrs.name ?? '';
    // r:id 经 xmlLite 去前缀后为 id
    const sheetPath = resolvePath('xl', workbookRels[sheetEl.attrs.id ?? ''] ?? '');
    if (!sheetPath || !zip.file(sheetPath)) continue;
    out.push({ sheetName, sheetPath, sheetRelsPath: relsPathOf(sheetPath) });
  }
  return out;
}

/** 'A1' / 'BC20' → 0 基行列；无法解析时返回 null */
export function decodeCellRef(ref: string): { row: number; col: number } | null {
  const m = ref.match(/^([A-Za-z]+)(\d+)/);
  if (!m) return null;
  const col = m[1].toUpperCase().split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const row = Number(m[2]) - 1;
  return col >= 0 && row >= 0 ? { row, col } : null;
}
