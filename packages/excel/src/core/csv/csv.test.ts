import { describe, expect, it } from 'vitest';
import { csvToSnapshot, displayValuesToRows, parseCsv, serializeCsv, trimEmptyEdges } from './csv';

describe('parseCsv', () => {
  it('解析 BOM、CRLF、引号转义与字段内换行', () => {
    const text = '\uFEFF名称,备注\r\n"甲,乙","他说""你好"""\r\n"多\n行",\r\n';
    expect(parseCsv(text)).toEqual([
      ['名称', '备注'],
      ['甲,乙', '他说"你好"'],
      ['多\n行', ''],
    ]);
  });

  it('兼容 LF、空行与末尾无换行', () => {
    expect(parseCsv('a,b\n\nc,d')).toEqual([['a', 'b'], [''], ['c', 'd']]);
    expect(parseCsv('')).toEqual([]);
  });

  it('引号包裹的空值与普通空字段都保留为字符串', () => {
    expect(parseCsv('"",x,')).toEqual([['', 'x', '']]);
  });
});

describe('serializeCsv', () => {
  it('输出 BOM + CRLF，并仅对必要字段加引号', () => {
    expect(serializeCsv([['a', 'b,c'], ['quote"here', 'line\nbreak'], [' padded ', '']])).toBe(
      '\uFEFFa,"b,c"\r\n"quote""here","line\nbreak"\r\n" padded ",',
    );
  });

  it('解析后再序列化保持数据一致', () => {
    const rows = [['名称', '备注'], ['甲,乙', '他说"你好"'], ['多\n行', '']];
    expect(parseCsv(serializeCsv(rows))).toEqual(rows);
  });
});

describe('trimEmptyEdges', () => {
  it('去掉末尾全空行与右侧全空列', () => {
    expect(trimEmptyEdges([['a', '', ''], ['', '', ''], ['', '', '']])).toEqual([['a']]);
    expect(trimEmptyEdges([['', ''], ['', '']])).toEqual([]);
  });
});

describe('csvToSnapshot', () => {
  it('生成单表快照：值全部按文本保存，前导零不丢', () => {
    const snap = csvToSnapshot('\uFEFF编号,金额\r\n0012,3.14\r\n', '销售');
    const sheet = snap.sheets['sheet-01']!;
    expect(sheet.name).toBe('销售');
    expect(snap.name).toBe('销售');
    expect(sheet.cellData[0]?.[0]).toEqual({ v: '编号', t: 1 });
    expect(sheet.cellData[1]?.[0]).toEqual({ v: '0012', t: 1 });
    expect(sheet.cellData[1]?.[1]).toEqual({ v: '3.14', t: 1 });
    expect(sheet.rowCount).toBeGreaterThanOrEqual(100);
    expect(sheet.columnCount).toBeGreaterThanOrEqual(20);
  });

  it('以公式符号开头的 CSV 值按文本处理，不执行公式', () => {
    const snap = csvToSnapshot('=SUM(A1:A2)', '公式文本');
    expect(snap.sheets['sheet-01']!.cellData[0]?.[0]).toEqual({ v: '=SUM(A1:A2)', t: 1 });
  });
});

describe('displayValuesToRows', () => {
  it('转换为字符串并裁剪末尾空行列', () => {
    expect(displayValuesToRows([['1.00', ''], ['', '']])).toEqual([['1.00']]);
  });

  it('去掉富文本显示值的段末换行，保留单元格内换行', () => {
    expect(displayValuesToRows([['ABCDEF\r\n'], ['第一行\n第二行\r\n']])).toEqual([
      ['ABCDEF'],
      ['第一行\n第二行'],
    ]);
  });
});
