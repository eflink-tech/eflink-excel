import { describe, expect, it } from 'vitest';
import { createDocument, createEmptySnapshot, newId } from './spreadsheet';

describe('工作簿快照工厂', () => {
  it('createEmptySnapshot 默认单表 100 行 × 20 列', () => {
    const snap = createEmptySnapshot();
    expect(snap.name).toBe('未命名表格');
    expect(snap.sheetOrder).toEqual(['sheet-01']);
    const sheet = snap.sheets['sheet-01'];
    expect(sheet.id).toBe('sheet-01');
    expect(sheet.name).toBe('Sheet1');
    expect(sheet.rowCount).toBe(100);
    expect(sheet.columnCount).toBe(20);
    expect(sheet.cellData).toEqual({});
  });

  it('createEmptySnapshot 支持自定义名称，且 id 唯一', () => {
    const a = createEmptySnapshot('我的表');
    const b = createEmptySnapshot('我的表');
    expect(a.name).toBe('我的表');
    expect(a.id).not.toBe(b.id);
    expect(a.id).toMatch(/^workbook-/);
  });
});

describe('文档工厂', () => {
  it('createDocument 生成 doc- 前缀 id，标题与快照名称一致', () => {
    const doc = createDocument('库存表');
    expect(doc.id).toMatch(/^doc-/);
    expect(doc.title).toBe('库存表');
    expect(doc.snapshot.name).toBe('库存表');
    expect(doc.createdAt).toBeGreaterThan(0);
    expect(doc.updatedAt).toBe(doc.createdAt);
  });

  it('newId 连续生成不重复', () => {
    const ids = new Set(Array.from({ length: 50 }, () => newId()));
    expect(ids.size).toBe(50);
  });
});
