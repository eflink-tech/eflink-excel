import { describe, expect, it } from 'vitest';
import * as ExcelJS from 'exceljs';
import { snapshotToWorkbook } from './snapshotToWorkbook';
import { workbookToSnapshot } from './workbookToSnapshot';
import type { WorkbookSnapshot } from '../../types/spreadsheet';

function sampleSnapshot(): WorkbookSnapshot {
  return {
    id: 'wb-test',
    name: '销售统计',
    sheetOrder: ['s1', 's2'],
    sheets: {
      s1: {
        id: 's1',
        name: '销售统计',
        rowCount: 20,
        columnCount: 8,
        cellData: {
          0: {
            0: { v: '月份', s: { bl: 1, fs: 12, cl: { rgb: '#FFFFFF' }, bg: { rgb: '#4472C4' } } },
            1: { v: '销售额', s: { bl: 1 } },
          },
          1: { 0: { v: '1月' }, 1: { v: 100, s: { n: { pattern: '#,##0.00' } } } },
          3: { 0: { v: '3月' }, 1: { v: 360 } },
          4: {
            0: { v: '合计', s: { bl: 1 } },
            1: { f: '=SUM(B2:B4)', v: 660, s: { bl: 1, n: { pattern: '#,##0.00' } } },
          },
          5: { 0: { v: '占比', s: { it: 1 } }, 1: { f: '=B4/B5', v: 0.55, s: { n: { pattern: '0.0%' } } } },
        },
        mergeData: [{ startRow: 6, startColumn: 1, endRow: 7, endColumn: 3 }],
        rowData: { 0: { h: 32 } },
        columnData: { 0: { w: 120 } },
      },
      s2: {
        id: 's2',
        name: '说明',
        rowCount: 10,
        columnCount: 5,
        cellData: { 0: { 0: { v: '备注页' } } },
      },
    },
  };
}

function roundTrip(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const wb = new ExcelJS.Workbook();
  snapshotToWorkbook(snapshot, wb);
  return workbookToSnapshot(wb, snapshot.name);
}

describe('xlsx 双向转换', () => {
  it('导出方向：快照写入 exceljs 工作簿', () => {
    const wb = new ExcelJS.Workbook();
    snapshotToWorkbook(sampleSnapshot(), wb);
    expect(wb.worksheets.map((ws) => ws.name)).toEqual(['销售统计', '说明']);
    const ws = wb.worksheets[0];
    expect(ws.getCell('A1').value).toBe('月份');
    expect(ws.getCell('B5').value).toMatchObject({ formula: 'SUM(B2:B4)', result: 660 });
    expect(ws.getCell('B5').numFmt).toBe('#,##0.00');
    expect(ws.getCell('A1').font.bold).toBe(true);
    const fill = ws.getCell('A1').fill;
    expect(fill && 'pattern' in fill ? fill.fgColor?.argb : undefined).toBe('FF4472C4');
    expect(ws.getCell('B7').isMerged).toBe(true);
  });

  it('回路：值/公式/数字格式/样式/合并/行列尺寸 无损还原', () => {
    const out = roundTrip(sampleSnapshot());
    const s1 = Object.values(out.sheets)[0];
    const gd = (r: number, c: number) => s1.cellData[r]?.[c];
    // 回路输入是内联样式对象；运行时快照才可能有样式 id 字符串
    const st = (r: number, c: number) => {
      const style = gd(r, c)?.s;
      return typeof style === 'string' ? undefined : style;
    };

    expect(gd(0, 0)?.v).toBe('月份');
    expect(gd(1, 1)?.v).toBe(100);
    expect(gd(4, 1)?.f).toBe('=SUM(B2:B4)');
    expect(gd(4, 1)?.v).toBe(660);
    expect(st(5, 1)?.n?.pattern).toBe('0.0%');
    expect(st(0, 0)?.bl).toBe(1);
    expect(st(0, 0)?.fs).toBe(12);
    expect(st(0, 0)?.cl?.rgb?.toLowerCase()).toBe('#ffffff');
    expect(st(0, 0)?.bg?.rgb?.toLowerCase()).toBe('#4472c4');
    expect(st(5, 0)?.it).toBe(1);
    expect(s1.mergeData).toEqual([{ startRow: 6, startColumn: 1, endRow: 7, endColumn: 3 }]);
    expect(s1.rowData?.[0]?.h).toBe(32);
    expect(s1.columnData?.[0]?.w).toBe(120);
  });

  it('多工作表顺序与名称保持', () => {
    const out = roundTrip(sampleSnapshot());
    const sheets = out.sheetOrder?.map((id) => out.sheets[id]);
    expect(sheets?.map((s) => s.name)).toEqual(['销售统计', '说明']);
  });

  it('空快照导出导入不报错', () => {
    const empty: WorkbookSnapshot = {
      id: 'wb-empty',
      name: '空白',
      sheets: { 'sheet-01': { id: 'sheet-01', name: 'Sheet1', rowCount: 50, columnCount: 20, cellData: {} } },
    };
    const out = roundTrip(empty);
    expect(Object.keys(out.sheets)).toHaveLength(1);
    expect(out.sheets['sheet-01'].cellData).toEqual({});
  });

  it('Univer 运行时快照：cell.s 为样式 id 时经 workbook.styles 解析导出', () => {
    const runtimeSnapshot: WorkbookSnapshot = {
      id: 'wb-runtime',
      name: '运行时',
      styles: { s1: { bl: 1, fs: 14, cl: { rgb: '#FF0000' } }, s2: { n: { pattern: '0.00%' } } },
      sheets: {
        'sheet-01': {
          id: 'sheet-01',
          name: 'Sheet1',
          rowCount: 10,
          columnCount: 5,
          cellData: {
            0: { 0: { v: '标题', s: 's1' }, 1: { v: 0.95, s: 's2' }, 2: { f: '=1+2', v: 3 } },
          },
        },
      },
    };
    const wb = new ExcelJS.Workbook();
    snapshotToWorkbook(runtimeSnapshot, wb);
    const a1 = wb.worksheets[0].getCell('A1');
    expect(a1.font.bold).toBe(true);
    expect(a1.font.size).toBe(14);
    expect(a1.font.color?.argb).toBe('FFFF0000');
    expect(wb.worksheets[0].getCell('B1').numFmt).toBe('0.00%');
    expect(wb.worksheets[0].getCell('C1').value).toMatchObject({ formula: '1+2', result: 3 });
  });
});
