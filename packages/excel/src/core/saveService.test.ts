import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

// 局部 mock 默认存储：默认透传真实内存实现，可注入写库失败
let failUpdate = false;
vi.mock('../storage/registry', async () => {
  const { memoryStorage } = await import('../storage/memory');
  const base = memoryStorage();
  return {
    getDefaultStorage: () => ({
      ...base,
      async updateContent(...args: Parameters<typeof base.updateContent>) {
        if (failUpdate) throw new Error('mock 写库失败');
        return base.updateContent(...args);
      },
    }),
    setDefaultStorage: () => {},
  };
});

import {
  bindEditor,
  currentSnapshot,
  notifyChanged,
  saveNow,
  unbindEditor,
} from './saveService';
import { getDefaultStorage } from '../storage/registry';
import { useEditorStore } from '../store/editorStore';
import { createDocument } from '../types/spreadsheet';
import type { WorkbookSnapshot } from '../types/spreadsheet';

beforeEach(async () => {
  failUpdate = false;
  const storage = getDefaultStorage();
  await storage.clear?.();
  const doc = createDocument('保存服务测试');
  await storage.save(doc);
  useEditorStore.getState().openDoc({ id: doc.id, title: doc.title });
});

afterEach(() => {
  unbindEditor();
  vi.useRealTimers();
});

function boundSnapshot(): WorkbookSnapshot {
  const doc = createDocument('占位');
  doc.snapshot.sheets['sheet-01'].cellData[0] = { 0: { v: '当前内容' } };
  return doc.snapshot;
}

describe('saveService 保存服务', () => {
  it('未绑定时 notifyChanged / currentSnapshot 安全无副作用', () => {
    expect(currentSnapshot()).toBeNull();
    expect(() => notifyChanged()).not.toThrow();
    expect(useEditorStore.getState().dirty).toBe(false);
  });

  // 注意：saveService 的 autosaver 固定 1500ms，且 fake-indexeddb 依赖真实定时器，
  // 这里用真实计时 + 略长于防抖周期的等待来验证防抖落库。
  it('notifyChanged 标记脏状态，防抖到期后自动落库', async () => {
    let latest: WorkbookSnapshot | null = null;
    bindEditor(useEditorStore.getState().docId!, () => latest);

    notifyChanged();
    expect(useEditorStore.getState().dirty).toBe(true);

    latest = boundSnapshot();
    await new Promise((r) => setTimeout(r, 1800));

    const saved = await getDefaultStorage().load(useEditorStore.getState().docId!);
    expect(saved?.snapshot.sheets['sheet-01'].cellData[0]?.[0]?.v).toBe('当前内容');
    expect(useEditorStore.getState().dirty).toBe(false);
    expect(useEditorStore.getState().savedAt).not.toBeNull();
  }, 10_000);

  it('saveNow 立即落库最新标题与快照', async () => {
    let latest: WorkbookSnapshot | null = boundSnapshot();
    bindEditor(useEditorStore.getState().docId!, () => latest);

    useEditorStore.getState().setTitle('手动保存的标题');
    await saveNow();

    const saved = await getDefaultStorage().load(useEditorStore.getState().docId!);
    expect(saved?.title).toBe('手动保存的标题');
    expect(saved?.snapshot.sheets['sheet-01'].cellData[0]?.[0]?.v).toBe('当前内容');
    expect(useEditorStore.getState().saving).toBe(false);
  });

  it('currentSnapshot 返回绑定 getter 的最新值', () => {
    let latest: WorkbookSnapshot | null = null;
    bindEditor('doc-x', () => latest);
    expect(currentSnapshot()).toBeNull();
    latest = boundSnapshot();
    expect(currentSnapshot()?.sheets['sheet-01'].cellData[0]?.[0]?.v).toBe('当前内容');
  });

  it('getSnapshot 返回 null 时不落库也不报错', async () => {
    let latest: WorkbookSnapshot | null = null;
    bindEditor(useEditorStore.getState().docId!, () => latest);
    await expect(saveNow()).resolves.toBeUndefined();
  });

  it('写库失败时 saveNow 抛错且 saving 复位，不误报已保存', async () => {
    failUpdate = true;
    let latest: WorkbookSnapshot | null = boundSnapshot();
    bindEditor(useEditorStore.getState().docId!, () => latest);

    await expect(saveNow()).rejects.toThrow('mock 写库失败');
    expect(useEditorStore.getState().saving).toBe(false);
    expect(useEditorStore.getState().savedAt).toBeNull();
    expect(useEditorStore.getState().dirty).toBe(false);
  });

  it('unbindEditor 取消待执行的自动保存', async () => {
    let latest: WorkbookSnapshot | null = boundSnapshot();
    const docId = useEditorStore.getState().docId!;
    bindEditor(docId, () => latest);

    notifyChanged();
    unbindEditor();
    await new Promise((r) => setTimeout(r, 1800));

    const saved = await getDefaultStorage().load(docId);
    expect(saved?.snapshot.sheets['sheet-01'].cellData[0]).toBeUndefined();
  }, 10_000);
});
