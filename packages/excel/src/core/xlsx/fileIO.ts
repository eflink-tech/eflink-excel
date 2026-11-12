// 浏览器侧文件 IO：xlsx 导入导出、PNG 导出、下载工具
// exceljs / html2canvas 体量大且只在导入导出/截图时用到，改为调用点动态加载，
// 避免编辑器首屏就把它们打进同一个 chunk。
import JSZip from 'jszip';
import { snapshotToWorkbook } from './snapshotToWorkbook';
import { workbookToSnapshot } from './workbookToSnapshot';
import { clampOversizedDataValidations } from './dataValidationSanitize';
import { findAll, findFirst, parseXml } from './xmlLite';
import { extractCharts, type ImportedChart } from './chartImport';
import { extractAnnotations, type ImportedAnnotations } from './annotationImport';
import { extractRetentions, type ImportedRetentions } from './retentionImport';
import { buildChartOption, DEFAULT_CHART_ELEMENTS, renderChartFrameToDataUrl } from '../univer/chartOptions';
import type { WorkbookSnapshot } from '../../types/spreadsheet';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export interface ImportXlsxResult {
  snapshot: WorkbookSnapshot;
  /** 渲染成图片的待插入图表（编辑器挂载后落位，见 univer/importedCharts.ts） */
  charts: ImportedChart[];
  /** 待落位的超链接/批注（编辑器挂载后经 facade 写入，见 univer/importedAnnotations.ts） */
  annotations: ImportedAnnotations;
  /** 待落位的条件格式/数据验证（编辑器挂载后经 facade/命令写入，见 univer/importedRetentions.ts） */
  retentions: ImportedRetentions;
}

export async function exportXlsx(snapshot: WorkbookSnapshot, title: string): Promise<void> {
  const ExcelJS = await import('exceljs');
  const wb = new ExcelJS.Workbook();
  snapshotToWorkbook(snapshot, wb);
  const buffer = await wb.xlsx.writeBuffer();
  downloadBlob(new Blob([buffer], { type: XLSX_MIME }), `${title}.xlsx`);
}

export async function importXlsx(file: File): Promise<ImportXlsxResult> {
  const buffer = await file.arrayBuffer();
  const ExcelJS = await import('exceljs');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await clampOversizedDataValidations(buffer));
  const normalFontSize = await readNormalFontSize(buffer).catch(() => undefined);
  const snapshot = workbookToSnapshot(wb, trimExt(file.name), { normalFontSize });
  const [charts, annotations, retentions] = await Promise.all([
    renderImportedCharts(await extractCharts(buffer).catch(() => [])),
    extractAnnotations(buffer).catch(() => ({ links: [], comments: [] })),
    extractRetentions(buffer).catch(() => ({ conditionalFormats: [], dataValidations: [] })),
  ]);
  return { snapshot, charts, annotations, retentions };
}

/** 只解析工作簿快照（不渲染图表/批注），供导入预览使用 */
export async function parseXlsxSnapshot(file: File): Promise<WorkbookSnapshot> {
  const buffer = await file.arrayBuffer();
  const ExcelJS = await import('exceljs');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await clampOversizedDataValidations(buffer));
  const normalFontSize = await readNormalFontSize(buffer).catch(() => undefined);
  return workbookToSnapshot(wb, trimExt(file.name), { normalFontSize });
}

/** styles.xml → Normal 字号：cellStyleXfs 首个 xf 的 fontId 指向字体（缺省 0），列宽字符单位以此换算 */
async function readNormalFontSize(buffer: ArrayBuffer): Promise<number | undefined> {
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file('xl/styles.xml')?.async('string');
  if (!xml) return undefined;
  const root = parseXml(xml);
  const fonts = findAll(root, 'font');
  const styleXfs = findFirst(root, 'cellStyleXfs');
  const firstXf = findAll(styleXfs, 'xf')[0];
  const font = fonts[Number(firstXf?.attrs.fontId ?? 0)] ?? fonts[0];
  const size = Number(findAll(font, 'sz')[0]?.attrs.val);
  return Number.isFinite(size) && size > 0 ? size : undefined;
}

/** 解析出的图表渲染为 PNG dataUrl（单张失败跳过，不影响其余与单元格导入） */
async function renderImportedCharts(charts: ImportedChart[]): Promise<ImportedChart[]> {
  const rendered: ImportedChart[] = [];
  for (const chart of charts) {
    try {
      const option = buildChartOption([], chart.typeId, false, {
        elements: { ...DEFAULT_CHART_ELEMENTS, chartTitle: Boolean(chart.title) },
        title: chart.title,
      }, chart.explicit);
      chart.dataUrl = await renderChartFrameToDataUrl(option);
      rendered.push(chart);
    } catch (err) {
      console.warn('[chartImport] 图表渲染失败，已跳过', chart.sheetName, chart.rangeRef, err);
    }
  }
  return rendered;
}

/** 截取编辑器可视区域为 PNG */
export async function exportPng(container: HTMLElement, title: string): Promise<void> {
  const { default: html2canvas } = await import('html2canvas');
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

export function trimExt(filename: string): string {
  return filename.replace(/\.(xlsx|xlsm|xls|csv)$/i, '');
}
