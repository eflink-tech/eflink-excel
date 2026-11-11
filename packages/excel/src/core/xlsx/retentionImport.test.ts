// 条件格式 / 数据验证导入解析测试：jszip 构造最小 xlsx 包，覆盖规则类型、dxf 样式、
// 优先级排序、sqref 多区域、showDropDown 语义翻转，以及待落位队列。
import { beforeEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
  extractRetentions,
  peekImportedRetention,
  queueImportedRetentions,
  shiftImportedRetention,
  unshiftImportedRetention,
} from './retentionImport';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const RELS_NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';

async function buildXlsx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0"?><workbook ${NS}><sheets><sheet name="数据" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0"?><Relationships ${RELS_NS}><Relationship Id="rId1" Type=".../worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  );
  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0"?><styleSheet ${NS}><dxfs count="3">` +
      `<dxf><font><b/><color rgb="FFFF0000"/></font><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf>` +
      `<dxf><font><i/></font></dxf>` +
      `<dxf><font><u/><strike/></font></dxf>` +
      `</dxfs></styleSheet>`,
  );
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0"?><worksheet ${NS}>` +
      `<conditionalFormatting sqref="A1:A10 C1:C10">` +
      `<cfRule type="cellIs" operator="between" priority="2" dxfId="0"><formula>10</formula><formula>20</formula></cfRule>` +
      `</conditionalFormatting>` +
      `<conditionalFormatting sqref="B1:B10">` +
      `<cfRule type="containsText" operator="containsText" text="错误" priority="1" dxfId="1"><formula>NOT(ISERROR(SEARCH("错误",B1)))</formula></cfRule>` +
      `<cfRule type="duplicateValues" priority="3" dxfId="2"/>` +
      `<cfRule type="aboveAverage" aboveAverage="0" priority="4"/>` +
      `<cfRule type="top10" rank="5" bottom="1" percent="1" priority="5"/>` +
      `</conditionalFormatting>` +
      `<conditionalFormatting sqref="D1:D10">` +
      `<cfRule type="colorScale" priority="6"><colorScale>` +
      `<cfvo type="min"/><cfvo type="max"/>` +
      `<color rgb="FFF8696B"/><color rgb="FFFFEB84"/>` +
      `</colorScale></cfRule>` +
      `</conditionalFormatting>` +
      `<conditionalFormatting sqref="E1:E10">` +
      `<cfRule type="dataBar" priority="7"><dataBar>` +
      `<cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/>` +
      `</dataBar></cfRule>` +
      `</conditionalFormatting>` +
      `<conditionalFormatting sqref="F1:F10">` +
      `<cfRule type="iconSet" priority="8"><iconSet iconSet="3TrafficLights1">` +
      `<cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>` +
      `</iconSet></cfRule>` +
      `</conditionalFormatting>` +
      `<dataValidations count="2">` +
      `<dataValidation type="list" allowBlank="1" showDropDown="1" sqref="G1:G5"><formula1>"是,否"</formula1></dataValidation>` +
      `<dataValidation type="whole" operator="between" sqref="H1:H9" showErrorMessage="1" errorStyle="stop" ` +
      `errorTitle="范围" error="请输入 1-9" showInputMessage="1" promptTitle="提示" prompt="输入数字">` +
      `<formula1>1</formula1><formula2>9</formula2></dataValidation>` +
      `</dataValidations>` +
      `</worksheet>`,
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

describe('retentionImport 条件格式 / 数据验证解析', () => {
  it('条件格式：多区域 sqref、优先级排序、dxf 字体与填充色', async () => {
    const { conditionalFormats } = await extractRetentions(await buildXlsx());
    expect(conditionalFormats).toHaveLength(8);

    // 优先级按升序，containsText(1) 排在 between(2) 之前
    expect(conditionalFormats[0]).toMatchObject({ sheetName: '数据', priority: 1, stopIfTrue: false });
    expect(conditionalFormats[0].ranges).toEqual([{ startRow: 0, startColumn: 1, endRow: 9, endColumn: 1 }]);
    expect(conditionalFormats[0].rule).toMatchObject({
      type: 'highlightCell',
      subType: 'text',
      operator: 'containsText',
      value: '错误',
      style: { it: 1 },
    });

    const between = conditionalFormats.find((r) => (r.rule as { subType?: string }).subType === 'number')!;
    expect(between.ranges).toEqual([
      { startRow: 0, startColumn: 0, endRow: 9, endColumn: 0 },
      { startRow: 0, startColumn: 2, endRow: 9, endColumn: 2 },
    ]);
    expect(between.rule).toMatchObject({
      type: 'highlightCell',
      subType: 'number',
      operator: 'between',
      value: [10, 20],
      style: { bl: 1, cl: { rgb: '#ff0000' }, bg: { rgb: '#ffff00' } },
    });

    const duplicate = conditionalFormats.find((r) => (r.rule as { subType?: string }).subType === 'duplicateValues')!;
    expect(duplicate.rule).toMatchObject({ style: { ul: { s: 1 }, st: { s: 1 } } });

    const average = conditionalFormats.find((r) => (r.rule as { subType?: string }).subType === 'average')!;
    expect(average.rule).toMatchObject({ operator: 'lessThan' });

    const rank = conditionalFormats.find((r) => (r.rule as { subType?: string }).subType === 'rank')!;
    expect(rank.rule).toMatchObject({ isBottom: true, isPercent: true, value: 5 });

    const colorScale = conditionalFormats.find((r) => r.rule.type === 'colorScale')!;
    expect(colorScale.rule).toMatchObject({
      config: [
        { index: 0, value: { type: 'min' }, color: '#f8696b' },
        { index: 1, value: { type: 'max' }, color: '#ffeb84' },
      ],
    });

    const dataBar = conditionalFormats.find((r) => r.rule.type === 'dataBar')!;
    expect(dataBar.rule).toMatchObject({
      isShowValue: true,
      config: { isGradient: true, positiveColor: '#638ec6', nativeColor: '#638ec6' },
    });

    const iconSet = conditionalFormats.find((r) => r.rule.type === 'iconSet')!;
    expect(iconSet.rule).toEqual({
      type: 'iconSet',
      isShowValue: true,
      iconType: '3TrafficLights1',
      values: [
        { type: 'percent', value: 0 },
        { type: 'percent', value: 33 },
        { type: 'percent', value: 67 },
      ],
    });
  });

  it('数据验证：类型 / 操作符 / 公式与 showDropDown 语义翻转', async () => {
    const { dataValidations } = await extractRetentions(await buildXlsx());
    expect(dataValidations).toHaveLength(2);

    const list = dataValidations[0];
    expect(list).toMatchObject({
      sheetName: '数据',
      type: 'list',
      allowBlank: true,
      showDropDown: false, // Excel showDropDown=1 表示隐藏下拉箭头
      formula1: '"是,否"',
    });
    expect(list.ranges).toEqual([{ startRow: 0, startColumn: 6, endRow: 4, endColumn: 6 }]);

    const whole = dataValidations[1];
    expect(whole).toMatchObject({
      type: 'whole',
      operator: 'between',
      formula1: '1',
      formula2: '9',
      showErrorMessage: true,
      errorStyle: 1,
      errorTitle: '范围',
      error: '请输入 1-9',
      showInputMessage: true,
      promptTitle: '提示',
      prompt: '输入数字',
    });
    expect(whole.showDropDown).toBeUndefined();
  });

  it('不落位未知类型 / 无效区域', async () => {
    const zip = new JSZip();
    zip.file(
      'xl/workbook.xml',
      `<?xml version="1.0"?><workbook ${NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    );
    zip.file(
      'xl/_rels/workbook.xml.rels',
      `<?xml version="1.0"?><Relationships ${RELS_NS}><Relationship Id="rId1" Type=".../worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    );
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<?xml version="1.0"?><worksheet ${NS}>` +
        `<conditionalFormatting sqref="not-a-ref"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting>` +
        `<conditionalFormatting sqref="A1"><cfRule type="expression" priority="2"><formula>A1&gt;0</formula></cfRule></conditionalFormatting>` +
        `<dataValidations><dataValidation type="none" sqref="B1"/><dataValidation type="unknown" sqref="B1"/></dataValidations>` +
        `</worksheet>`,
    );
    const result = await extractRetentions(await zip.generateAsync({ type: 'arraybuffer' }));
    expect(result.conditionalFormats).toHaveLength(1);
    expect(result.dataValidations).toHaveLength(0);
  });
});

describe('retentionImport 待落位队列', () => {
  beforeEach(() => {
    // 清空队列（模块级状态，测试间隔离）
    while (shiftImportedRetention()) {
      /* drain */
    }
  });

  it('按条件格式 → 数据验证顺序排队，支持 peek / shift / unshift', () => {
    queueImportedRetentions({
      conditionalFormats: [
        { sheetName: 'S', ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }], priority: 1, stopIfTrue: false, rule: { type: 'colorScale', config: [] } as never },
      ],
      dataValidations: [
        { sheetName: 'S', ranges: [{ startRow: 1, startColumn: 0, endRow: 1, endColumn: 0 }], type: 'list' },
      ],
    });

    expect(peekImportedRetention()?.kind).toBe('cf');
    expect(shiftImportedRetention()?.kind).toBe('cf');
    expect(peekImportedRetention()?.kind).toBe('dv');
    const dv = shiftImportedRetention();
    if (!dv || dv.kind !== 'dv') throw new Error('expected dv task');
    expect(dv.type).toBe('list');

    unshiftImportedRetention(dv);
    expect(peekImportedRetention()).toEqual(dv);
    expect(shiftImportedRetention()).toEqual(dv);
    expect(peekImportedRetention()).toBeNull();
  });

  it('重新排队会替换旧队列（一次导入一份结果）', () => {
    queueImportedRetentions({ conditionalFormats: [], dataValidations: [{ sheetName: 'A', ranges: [], type: 'list' }] });
    queueImportedRetentions({ conditionalFormats: [], dataValidations: [{ sheetName: 'B', ranges: [], type: 'whole' }] });
    const task = peekImportedRetention();
    if (!task || task.kind !== 'dv') throw new Error('expected dv task');
    expect(task.sheetName).toBe('B');
  });
});
