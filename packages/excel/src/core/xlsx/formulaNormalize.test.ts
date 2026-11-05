import { describe, expect, it } from 'vitest';
import { normalizeSnapshotFormulas } from './formulaNormalize';
import type { WorkbookSnapshot } from '../../types/spreadsheet';

function snapshotOf(formula: string): WorkbookSnapshot {
  return {
    id: 'wb',
    name: 't',
    sheetOrder: ['s1'],
    sheets: {
      s1: {
        id: 's1',
        name: 'Sheet1',
        rowCount: 10,
        columnCount: 5,
        cellData: { 0: { 0: { f: formula, v: 1 }, 1: { v: '文本' } } },
      },
    },
  };
}

describe('normalizeSnapshotFormulas', () => {
  it('去掉历史双等号公式', () => {
    const snap = snapshotOf("==SUM('①-功能需求'!BO66+'①-功能需求'!BZ66)");
    expect(normalizeSnapshotFormulas(snap)).toBe(true);
    expect(snap.sheets.s1.cellData[0]![0]!.f).toBe("=SUM('①-功能需求'!BO66+'①-功能需求'!BZ66)");
  });

  it('正常单等号公式与非公式单元格不变', () => {
    const snap = snapshotOf('=SUM(A1:A2)');
    expect(normalizeSnapshotFormulas(snap)).toBe(false);
    expect(snap.sheets.s1.cellData[0]![0]!.f).toBe('=SUM(A1:A2)');
    expect(snap.sheets.s1.cellData[0]![1]!.v).toBe('文本');
  });
});
