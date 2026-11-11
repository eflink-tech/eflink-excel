import { describe, expect, it } from 'vitest';
import { normalizeSnapshotCellTypes, normalizeSnapshotFormulas } from './formulaNormalize';
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

describe('normalizeSnapshotCellTypes', () => {
  it('按 v 的类型补全缺失的 t，已有 t 不覆盖', () => {
    const snap: WorkbookSnapshot = {
      id: 'wb',
      name: 't',
      sheetOrder: ['s1'],
      sheets: {
        s1: {
          id: 's1',
          name: 'Sheet1',
          rowCount: 3,
          columnCount: 3,
          cellData: {
            0: {
              0: { v: 2 },
              1: { v: '文本' },
              2: { v: true },
            },
            1: {
              0: { v: 5, t: 1 }, // 强制文本：保留原 t
              1: { f: '=A1+A2', v: 4 }, // 公式缓存结果按结果类型
              2: {}, // 空单元格不动
            },
          },
        },
      },
    };
    expect(normalizeSnapshotCellTypes(snap)).toBe(true);
    expect(snap.sheets.s1.cellData[0]![0]!.t).toBe(2);
    expect(snap.sheets.s1.cellData[0]![1]!.t).toBe(1);
    expect(snap.sheets.s1.cellData[0]![2]!.t).toBe(3);
    expect(snap.sheets.s1.cellData[1]![0]!.t).toBe(1);
    expect(snap.sheets.s1.cellData[1]![1]!.t).toBe(2);
    expect(snap.sheets.s1.cellData[1]![2]!.t).toBeUndefined();
    expect(normalizeSnapshotCellTypes(snap)).toBe(false);
  });
});
