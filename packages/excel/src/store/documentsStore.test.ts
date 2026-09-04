import { beforeEach, describe, expect, it } from 'vitest';
import { memoryStorage } from '../storage/memory';
import { setDefaultStorage } from '../storage/registry';
import { useDocumentsStore } from './documentsStore';

beforeEach(async () => {
  setDefaultStorage(memoryStorage());
  useDocumentsStore.setState({ docs: [] });
});

describe('documentsStore 文档列表管理', () => {
  it('createDoc 落库并刷新列表', async () => {
    const doc = await useDocumentsStore.getState().createDoc('测试文档');
    expect(doc.title).toBe('测试文档');

    const { docs } = useDocumentsStore.getState();
    expect(docs).toHaveLength(1);
    expect(docs[0]).toMatchObject({ id: doc.id, title: '测试文档' });
  });

  it('addDoc 导入场景直接落库完整文档', async () => {
    const doc = await useDocumentsStore.getState().createDoc('草稿');
    doc.title = '导入的表格';
    doc.snapshot.sheets['sheet-01'].cellData[0] = { 0: { v: '外部数据' } };
    await useDocumentsStore.getState().addDoc(doc);

    const { docs } = useDocumentsStore.getState();
    expect(docs.map((d) => d.title)).toContain('导入的表格');
  });

  it('removeDoc 删除后列表同步刷新', async () => {
    const a = await useDocumentsStore.getState().createDoc('A');
    await useDocumentsStore.getState().createDoc('B');
    await useDocumentsStore.getState().removeDoc(a.id);

    const { docs } = useDocumentsStore.getState();
    expect(docs.map((d) => d.title)).toEqual(['B']);
  });

  it('renameDoc 修改标题并提升更新时间', async () => {
    const doc = await useDocumentsStore.getState().createDoc('旧名');
    await new Promise((r) => setTimeout(r, 5));
    await useDocumentsStore.getState().renameDoc(doc.id, '新名');

    const { docs } = useDocumentsStore.getState();
    expect(docs[0]?.title).toBe('新名');
    expect(docs[0]?.updatedAt).toBeGreaterThanOrEqual(doc.updatedAt);
  });

  it('refresh 从数据库重新拉取列表', async () => {
    await useDocumentsStore.getState().createDoc('X');
    useDocumentsStore.setState({ docs: [] });
    expect(useDocumentsStore.getState().docs).toHaveLength(0);

    await useDocumentsStore.getState().refresh();
    expect(useDocumentsStore.getState().docs.map((d) => d.title)).toEqual(['X']);
  });
});
