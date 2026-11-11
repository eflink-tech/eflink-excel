// 打印 / PDF：按当前 sheet 已用区域重建可打印表格，经隐藏 iframe 调用系统打印。
// 浏览器打印对话框自带「另存为 PDF」，无需引入额外 PDF 库。
import { buildSheetImageElement, type SheetImageData } from '../xlsx/sheetImage';

export type PrintOrientation = 'portrait' | 'landscape';

export interface PrintOptions {
  orientation: PrintOrientation;
  title: string;
  sheetName: string;
}

/** A4 在 96dpi 下的可打印宽度（扣除 10mm 页边距），用于把宽表缩放到页宽 */
const PAGE_WIDTH_PX: Record<PrintOrientation, number> = {
  portrait: 688,
  landscape: 1046,
};

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** 生成独立打印文档：标题 + 表名 + 整表，宽表自动缩放到页宽 */
export function buildPrintHtml(data: SheetImageData, options: PrintOptions): string {
  const table = buildSheetImageElement(data);
  const tableWidth = Number.parseFloat(table.style.width) || 800;
  const scale = Math.min(1, PAGE_WIDTH_PX[options.orientation] / tableWidth);
  const size = options.orientation === 'landscape' ? 'landscape' : 'portrait';
  return `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(options.title)}</title>
<style>
  @page { size: A4 ${size}; margin: 10mm; }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; background: #fff; }
  body { color: #1f2329; font-family: Arial, 'PingFang SC', 'Microsoft YaHei', sans-serif; }
  .print-head {
    display: flex; align-items: baseline; justify-content: space-between; gap: 12px;
    margin: 0 0 6px; padding-bottom: 4px; border-bottom: 1px solid #d0d5dd;
  }
  .print-title { font-size: 12pt; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .print-sheet { font-size: 9pt; color: #667085; white-space: nowrap; }
  .print-body { zoom: ${scale}; }
  @media print {
    .print-head { break-inside: avoid; }
    tr, td { break-inside: avoid; }
  }
</style>
</head>
<body>
  <div class="print-head">
    <div class="print-title">${escapeHtml(options.title)}</div>
    <div class="print-sheet">${escapeHtml(options.sheetName)}</div>
  </div>
  <div class="print-body">${table.outerHTML}</div>
</body>
</html>`;
}

/** 经隐藏 iframe 打印（不打开新窗口，避免被弹窗拦截）；打印对话框里可另存为 PDF */
export function printHtml(html: string): void {
  const iframe = document.createElement('iframe');
  iframe.setAttribute('aria-hidden', 'true');
  iframe.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden;';
  iframe.onload = () => {
    try {
      iframe.contentWindow?.focus();
      iframe.contentWindow?.print();
    } finally {
      // 部分浏览器打印对话框会阻塞；延迟释放，避免打印尚未开始就销毁文档
      window.setTimeout(() => iframe.remove(), 60_000);
    }
  };
  iframe.srcdoc = html;
  document.body.appendChild(iframe);
}
