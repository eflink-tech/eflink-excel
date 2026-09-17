import { describe, expect, it } from 'vitest';
import * as ExcelJS from 'exceljs';
import { snapshotToWorkbook } from './snapshotToWorkbook';
import { workbookToSnapshot } from './workbookToSnapshot';
import type { CellStyle, WorkbookSnapshot } from '../../types/spreadsheet';

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

  it('富文本含换行时 dataStream 展开存储、run 区间按归一化文本计长', () => {
    const snap = importWorkbook((ws) => {
      ws.getCell('A1').value = {
        richText: [
          { text: 'A\nB', font: { bold: true } },
          { text: 'C' },
        ],
      };
    });
    const cell = snap.sheets['sheet-01'].cellData[0]?.[0];
    // dataStream 物理展开为 A\r\n（3 字符）；run 的 st/ed 按归一化文本（\r\n 记 1 字符）计长，C 在 [3, 4)
    expect(cell?.p?.body.dataStream).toBe('A\r\nBC\r\n');
    expect(cell?.p?.body.textRuns?.[0]?.ts?.bl).toBe(1);
  });

  it('富文本源含 \\r\\n 时归一化不双重展开', () => {
    const snap = importWorkbook((ws) => {
      ws.getCell('A1').value = {
        richText: [
          { text: 'A\r\nB', font: { bold: true } },
          { text: 'C' },
        ],
      };
    });
    const cell = snap.sheets['sheet-01'].cellData[0]?.[0];
    // 'A\r\nB' 保持 3 字符，不被展开成 \r\r\n；C 位于 [3, 4)
    expect(cell?.p?.body.dataStream).toBe('A\r\nBC\r\n');
    expect(cell?.p?.body.textRuns?.[0]).toMatchObject({ st: 0, ed: 3 });
  });

  it('冻结窗格仅 topLeftCell 时回退解析', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('回退表');
    ws.getCell('A1').value = '冻结区';
    ws.views = [{ state: 'frozen', topLeftCell: 'C4' }] as never;
    const snap = workbookToSnapshot(wb, '回退');
    expect(snap.sheets['sheet-01'].freeze).toEqual({ startRow: 3, startColumn: 2, xAxisSplit: 2, yAxisSplit: 3 });
  });
});

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
});
