import { beforeEach, describe, expect, it } from 'vitest';
import { indexedDbStorage } from './indexedDb';
import { createDocument } from '../types/spreadsheet';

const storage = indexedDbStorage();

beforeEach(async () => {
  await storage.clear?.();
});

describe('IndexedDB 存储适配器', () => {
  it('保存后可按 id 读回', async () => {
    const doc = createDocument('测试表');
    doc.snapshot.sheets['sheet-01'].cellData[0] = { 0: { v: 'hello' } };
    await storage.save(doc);

    const loaded = await storage.load(doc.id);
    expect(loaded?.title).toBe('测试表');
    expect(loaded?.snapshot.sheets['sheet-01'].cellData[0]?.[0]?.v).toBe('hello');
  });

  it('列表按更新时间倒序', async () => {
    const a = createDocument('A');
    await storage.save(a);
    const b = createDocument('B');
    await storage.save({ ...b, updatedAt: a.updatedAt + 1000 });

    const docs = await storage.list();
    expect(docs.map((d) => d.title)).toEqual(['B', 'A']);
  });

  it('重命名只改标题不碰快照', async () => {
    const doc = createDocument('旧名');
    doc.snapshot.sheets['sheet-01'].cellData[1] = { 1: { v: 42 } };
    await storage.save(doc);

    await storage.rename(doc.id, '新名');
    const loaded = await storage.load(doc.id);
    expect(loaded?.title).toBe('新名');
    expect(loaded?.snapshot.sheets['sheet-01'].cellData[1]?.[1]?.v).toBe(42);
  });

  it('编辑器保存更新标题与快照', async () => {
    const doc = createDocument('未命名表格');
    await storage.save(doc);

    doc.snapshot.sheets['sheet-01'].cellData[0] = { 0: { v: 'changed' } };
    await storage.updateContent(doc.id, '改过的表', doc.snapshot);

    const loaded = await storage.load(doc.id);
    expect(loaded?.title).toBe('改过的表');
    expect(loaded?.snapshot.sheets['sheet-01'].cellData[0]?.[0]?.v).toBe('changed');
  });

  it('删除文档', async () => {
    const doc = createDocument('待删');
    await storage.save(doc);
    await storage.delete(doc.id);
    expect(await storage.load(doc.id)).toBeUndefined();
    expect(await storage.list()).toHaveLength(0);
  });
});
