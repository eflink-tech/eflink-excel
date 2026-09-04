// .efexcel 文档备份格式：SheetDocument（标题 + 快照 + 元信息）的 JSON 序列化
import { createDocument } from '../types/spreadsheet';
import type { SheetDocument } from '../types/spreadsheet';
import { downloadBlob } from './xlsx/fileIO';

const EFEXCEL_EXT = '.efexcel';

/** 导出文件的来源信息（写入协议 source 字段，便于追溯文件由哪个应用产生） */
export interface EfexcelSource {
  app: string;
  url: string;
}

/** 默认来源：易飞表格 */
export const DEFAULT_EFEXCEL_SOURCE: EfexcelSource = {
  app: '易飞表格',
  url: 'https://eflink.tech/office/excel',
};

/** 文件名安全化：去除各平台非法字符与首尾空白/点，避免下载文件名异常 */
function sanitizeFilename(title: string): string {
  const cleaned = title
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
    .replace(/[\s.]+$/g, '')
    .trim();
  return cleaned || '未命名表格';
}

/** 导出文档为 .efexcel 备份文件（JSON，含 source 来源信息） */
export function exportEfexcel(doc: SheetDocument, source: EfexcelSource = DEFAULT_EFEXCEL_SOURCE): void {
  const payload = { ...doc, source };
  const json = JSON.stringify(payload, null, 2);
  downloadBlob(new Blob([json], { type: 'application/json' }), `${sanitizeFilename(doc.title)}${EFEXCEL_EXT}`);
}

/**
 * 解析 .efexcel / .json 备份文件；结构不合法时抛错。
 * 返回全新 id 的文档副本，避免导入后落库覆盖源文档。
 */
export async function importEfexcel(file: File): Promise<SheetDocument> {
  let raw: unknown;
  try {
    raw = JSON.parse(await file.text());
  } catch {
    throw new Error('文件不是合法 JSON');
  }
  const doc = raw as Partial<SheetDocument> | null;
  if (
    !doc?.snapshot ||
    typeof doc.snapshot.sheets !== 'object' ||
    doc.snapshot.sheets === null ||
    !Array.isArray(doc.snapshot.sheetOrder) ||
    doc.snapshot.sheetOrder.length === 0
  ) {
    throw new Error('缺少有效的表格快照');
  }
  const imported = createDocument(doc.title || '导入的表格');
  imported.snapshot = doc.snapshot;
  return imported;
}
