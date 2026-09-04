// 浏览器侧文件 IO：xlsx 导入导出、PNG 导出、下载工具
import * as ExcelJS from 'exceljs';
import html2canvas from 'html2canvas';
import { snapshotToWorkbook } from './snapshotToWorkbook';
import { workbookToSnapshot } from './workbookToSnapshot';
import type { WorkbookSnapshot } from '../../types/spreadsheet';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export async function exportXlsx(snapshot: WorkbookSnapshot, title: string): Promise<void> {
  const wb = new ExcelJS.Workbook();
  snapshotToWorkbook(snapshot, wb);
  const buffer = await wb.xlsx.writeBuffer();
  downloadBlob(new Blob([buffer], { type: XLSX_MIME }), `${title}.xlsx`);
}

export async function importXlsx(file: File): Promise<WorkbookSnapshot> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());
  return workbookToSnapshot(wb, trimExt(file.name));
}

/** 截取编辑器可视区域为 PNG */
export async function exportPng(container: HTMLElement, title: string): Promise<void> {
  const canvas = await html2canvas(container, { backgroundColor: '#ffffff', logging: false });
  await new Promise<void>((resolve) =>
    canvas.toBlob((blob) => {
      if (blob) downloadBlob(blob, `${title}.png`);
      resolve();
    }, 'image/png'),
  );
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  // Firefox 要求锚点在文档内 click 才可靠
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  a.remove();
  // 立即 revoke 会让部分浏览器（Safari 等）中断下载或丢失文件名，延后释放
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

function trimExt(filename: string): string {
  return filename.replace(/\.(xlsx|xlsm|xls)$/i, '');
}
