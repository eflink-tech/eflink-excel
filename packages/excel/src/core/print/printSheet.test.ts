// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { buildPrintHtml } from './printSheet';
import type { SheetImageData } from '../xlsx/sheetImage';

function makeData(widths: number[] = [100]): SheetImageData {
  return {
    sheet: { id: 's1', name: '销售表', rowCount: 10, columnCount: 10, cellData: { 0: { 0: { v: '标题' } } } },
    displayValues: [['标题', ...widths.slice(1).map(() => '')]],
    rawValues: [['标题', ...widths.slice(1).map(() => null)]],
    rowHeights: [24],
    columnWidths: widths,
  };
}

describe('buildPrintHtml', () => {
  it('生成 A4 纵向打印文档并转义标题', () => {
    const html = buildPrintHtml(makeData(), { orientation: 'portrait', title: '<季度> 报表', sheetName: '销售表' });
    expect(html).toContain('@page { size: A4 portrait; margin: 10mm; }');
    expect(html).toContain('&lt;季度&gt; 报表');
    expect(html).toContain('销售表');
    expect(html).toContain('<table');
    expect(html).toContain('zoom: 1;');
  });

  it('宽表按横向页宽缩放', () => {
    const html = buildPrintHtml(makeData([1000, 1000]), {
      orientation: 'landscape',
      title: '宽表',
      sheetName: 'Sheet1',
    });
    expect(html).toContain('@page { size: A4 landscape; margin: 10mm; }');
    expect(html).toMatch(/zoom: 0\.523;/);
  });
});
