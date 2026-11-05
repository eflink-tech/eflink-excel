// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { buildSheetImageElement, type SheetImageData } from './sheetImage';

function makeData(partial: Partial<SheetImageData> = {}): SheetImageData {
  return {
    sheet: {
      id: 's1',
      name: 'Sheet1',
      rowCount: 20,
      columnCount: 20,
      cellData: {},
    },
    displayValues: [['']],
    rawValues: [[null]],
    rowHeights: [24],
    columnWidths: [88],
    ...partial,
  };
}

describe('buildSheetImageElement 整表图片', () => {
  it('按已用区域生成全部行列，合并单元格跨列且只渲染左上角', () => {
    const data = makeData({
      sheet: {
        id: 's1',
        name: 'Sheet1',
        rowCount: 20,
        columnCount: 20,
        showGridlines: 0,
        mergeData: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }],
        cellData: {
          0: {
            0: { v: '标题', s: { bl: 1, ht: 2, bg: { rgb: '#ffeeaa' } } },
            2: { v: 42 },
          },
          1: { 0: { v: 'a' }, 1: { v: 'b' } },
        },
      },
      displayValues: [
        ['标题', '', '42'],
        ['a', 'b', ''],
      ],
      rawValues: [
        ['标题', null, 42],
        ['a', 'b', null],
      ],
      rowHeights: [30, 20],
      columnWidths: [100, 80, 60],
    });

    const table = buildSheetImageElement(data);
    expect(table.style.width).toBe('240px');

    const rows = table.querySelectorAll<HTMLTableRowElement>('tbody tr');
    expect(rows).toHaveLength(2);
    expect(rows[0].style.height).toBe('30px');

    const first = rows[0].querySelectorAll<HTMLTableCellElement>('td');
    // 第一行有 3 列，A1:B1 合并后剩 2 个 td
    expect(first).toHaveLength(2);
    expect(first[0].textContent).toBe('标题');
    expect(first[0].colSpan).toBe(2);
    expect(first[0].style.fontWeight).toBe('700');
    expect(first[0].style.textAlign).toBe('center');
    expect(first[0].style.background).not.toBe('');
    // 数字未显式对齐时默认右对齐
    expect(first[1].textContent).toBe('42');
    expect(first[1].style.textAlign).toBe('right');
    expect(first[1].style.borderTop).toBe(''); // 关闭网格线时无默认边框

    const cols = table.querySelectorAll('col');
    expect(cols).toHaveLength(3);
    expect((cols[2] as HTMLTableColElement).style.width).toBe('60px');
  });

  it('透传样式、自动换行与边框，富文本无显示值时回退正文', () => {
    const data = makeData({
      sheet: {
        id: 's1',
        name: 'Sheet1',
        rowCount: 20,
        columnCount: 20,
        cellData: {
          0: {
            0: {
              p: { id: 'p1', body: { dataStream: 'hello\r\nworld\r' }, documentStyle: {} },
              s: {
                fs: 14,
                cl: { rgb: '#00ff00' },
                tb: 3,
                ul: { s: 1 },
                bd: { t: { s: 1, cl: { rgb: '#ff0000' } } },
              },
            },
          },
        },
      },
      displayValues: [['']],
      rawValues: [[null]],
    });

    const td = buildSheetImageElement(data).querySelector('td') as HTMLTableCellElement;
    expect(td.textContent).toBe('hello\nworld');
    expect(td.style.whiteSpace).toBe('pre-wrap');
    expect(td.style.wordBreak).toBe('break-word');
    expect(td.style.fontSize).toBe('14pt');
    expect(td.style.color).toBe('rgb(0, 255, 0)');
    expect(td.style.textDecoration).toContain('underline');
    expect(td.style.borderTop).toBe('1px solid rgb(255, 0, 0)');
  });

  it('空内容也生成一个单元格，避免导出空白画布', () => {
    const table = buildSheetImageElement(makeData());
    expect(table.querySelectorAll('td')).toHaveLength(1);
  });
});
