// @vitest-environment node
import { describe, expect, it } from 'vitest';
import * as ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { clampOversizedDataValidations } from './dataValidationSanitize';

const SHEET_PATH = 'xl/worksheets/sheet1.xml';

async function workbookWithValidations(validationsXml: string): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('数据');
  ws.getCell('A1').value = '表头';
  ws.getCell('B3').value = 42;
  const buffer = (await wb.xlsx.writeBuffer()) as ArrayBuffer;

  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file(SHEET_PATH)!.async('string');
  zip.file(SHEET_PATH, xml.replace('</worksheet>', `${validationsXml}</worksheet>`));
  return zip.generateAsync({ type: 'arraybuffer' });
}

describe('clampOversizedDataValidations', () => {
  it('把整列验证收拢到已用行，丢弃完全在数据区之外的超大验证', async () => {
    const buffer = await workbookWithValidations(
      '<dataValidations count="2">' +
        '<dataValidation type="date" allowBlank="1" sqref="F2:F1048576"/>' +
        '<dataValidation type="list" allowBlank="1" sqref="M164:M1048576"/>' +
        '</dataValidations>',
    );

    const sanitized = await clampOversizedDataValidations(buffer);
    const zip = await JSZip.loadAsync(sanitized);
    const xml = await zip.file(SHEET_PATH)!.async('string');

    expect(xml).toContain('sqref="F2:F3"');
    expect(xml).not.toContain('M164');
    expect(xml).not.toContain('1048576');

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(sanitized);
    expect(wb.worksheets[0].getCell('A1').value).toBe('表头');
    const { model } = (wb.worksheets[0] as unknown as { dataValidations: { model: Record<string, unknown> } }).dataValidations;
    expect(Object.keys(model)).toEqual(['F2', 'F3']);
  });

  it('普通验证保持原样，且无需修改时返回原 buffer', async () => {
    const buffer = await workbookWithValidations(
      '<dataValidations count="1">' +
        '<dataValidation type="list" allowBlank="1" sqref="A1:A10"/>' +
        '</dataValidations>',
    );

    const sanitized = await clampOversizedDataValidations(buffer);
    expect(sanitized).toBe(buffer);
  });
});
