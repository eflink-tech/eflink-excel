// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import * as ExcelJS from 'exceljs';

// exportPng 依赖 html2canvas 渲染，jsdom 下用假画布代替
vi.mock('html2canvas', () => ({
  default: vi.fn(async () => {
    const canvas = document.createElement('canvas');
    canvas.width = 800;
    canvas.height = 600;
    canvas.getContext = () => null;
    canvas.toBlob = (cb: BlobCallback) => cb(new Blob(['fake-png'], { type: 'image/png' }));
    return canvas as unknown as HTMLCanvasElement;
  }),
}));

import { downloadBlob, exportPng, exportXlsx, importXlsx } from './fileIO';
import type { WorkbookSnapshot } from '../../types/spreadsheet';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function sampleSnapshot(): WorkbookSnapshot {
  return {
    id: 'wb-1',
    name: '导出测试',
    sheetOrder: ['s1'],
    sheets: {
      s1: {
        id: 's1',
        name: '导出测试',
        rowCount: 20,
        columnCount: 8,
        cellData: {
          0: { 0: { v: '标题', s: { bl: 1 } }, 1: { v: 42 } },
          1: { 0: { f: '=SUM(B1:B1)', v: 42 } },
        },
      },
    },
  };
}

/** jsdom 未实现 createObjectURL / 锚点导航，替换并记录每次下载（文件名 + blob） */
function captureDownload() {
  const downloads: { name: string; blob: Blob | null }[] = [];
  let pendingBlob: Blob | null = null;
  Object.defineProperty(URL, 'createObjectURL', {
    configurable: true,
    writable: true,
    value: vi.fn((blob: Blob) => {
      pendingBlob = blob;
      return 'blob:mock';
    }),
  });
  Object.defineProperty(URL, 'revokeObjectURL', { configurable: true, writable: true, value: vi.fn() });
  vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) {
    downloads.push({ name: this.download, blob: pendingBlob });
    pendingBlob = null;
  });
  return downloads;
}

describe('fileIO 浏览器侧文件 IO', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('importXlsx 解析 xlsx 文件为快照：文件名去扩展名、值/公式/样式还原', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('数据表');
    ws.getCell('A1').value = '表头';
    ws.getCell('A1').font = { bold: true, size: 14 };
    ws.getCell('B1').value = 100;
    ws.getCell('B1').numFmt = '#,##0.00';
    ws.getCell('A2').value = { formula: 'SUM(B1:B1)', result: 100 };
    ws.mergeCells('A3:C3');
    const buffer = await wb.xlsx.writeBuffer();

    const file = new File([buffer as ArrayBuffer], '年度报表.xlsx', { type: XLSX_MIME });
    const snapshot = await importXlsx(file);

    expect(snapshot.name).toBe('年度报表');
    const sheet = Object.values(snapshot.sheets)[0];
    // 导入方向产出的都是内联样式对象，此处收窄掉样式 id 字符串分支
    const st = (r: number, c: number) => {
      const style = sheet.cellData[r]?.[c]?.s;
      return typeof style === 'string' ? undefined : style;
    };
    expect(sheet.name).toBe('数据表');
    expect(sheet.cellData[0]?.[0]?.v).toBe('表头');
    expect(st(0, 0)?.bl).toBe(1);
    expect(st(0, 0)?.fs).toBe(14);
    expect(sheet.cellData[0]?.[1]?.v).toBe(100);
    expect(st(0, 1)?.n?.pattern).toBe('#,##0.00');
    expect(sheet.cellData[1]?.[0]?.f).toBe('=SUM(B1:B1)');
    expect(sheet.cellData[1]?.[0]?.v).toBe(100);
    expect(sheet.mergeData).toEqual([{ startRow: 2, startColumn: 0, endRow: 2, endColumn: 2 }]);
  });

  it('exportXlsx 生成 xlsx 下载：文件名带标题，内容可再解析', async () => {
    const downloads = captureDownload();

    await exportXlsx(sampleSnapshot(), '导出测试');

    expect(downloads).toHaveLength(1);
    expect(downloads[0].name).toBe('导出测试.xlsx');
    expect(downloads[0].blob?.type).toBe(XLSX_MIME);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await downloads[0].blob!.arrayBuffer());
    expect(wb.worksheets[0].getCell('A1').value).toBe('标题');
    expect(wb.worksheets[0].getCell('A1').font.bold).toBe(true);
    expect(wb.worksheets[0].getCell('B1').value).toBe(42);
  });

  it('downloadBlob 创建临时 url 并以指定文件名触发下载，延迟释放 url', () => {
    vi.useFakeTimers();
    try {
      const downloads = captureDownload();
      const blob = new Blob(['hello'], { type: 'text/plain' });

      downloadBlob(blob, '说明.txt');

      expect(downloads).toHaveLength(1);
      expect(downloads[0].name).toBe('说明.txt');
      expect(URL.createObjectURL).toHaveBeenCalledWith(blob);
      // url 延迟释放（避免部分浏览器下载被中断），同步阶段不 revoke
      expect(URL.revokeObjectURL).not.toHaveBeenCalled();
      vi.advanceTimersByTime(10_000);
      expect(URL.revokeObjectURL).toHaveBeenCalledWith('blob:mock');
    } finally {
      vi.useRealTimers();
    }
  });

  it('exportPng 用 html2canvas 截图并触发 png 下载', async () => {
    const downloads = captureDownload();
    const container = document.createElement('div');
    document.body.appendChild(container);

    await exportPng(container, '截图页');

    expect(downloads).toHaveLength(1);
    expect(downloads[0].name).toBe('截图页.png');
    expect(downloads[0].blob?.type).toBe('image/png');
  });

  it('importXlsx 对损坏文件抛出异常（由上层动作转为失败提示）', async () => {
    const bad = new File([new TextEncoder().encode('not an xlsx')], '坏文件.xlsx');
    await expect(importXlsx(bad)).rejects.toThrow();
  });
});
