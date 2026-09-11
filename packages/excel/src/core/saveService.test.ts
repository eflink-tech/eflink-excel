// @vitest-environment jsdom
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

const DRAFT_PREFIX = 'eflink:draft:excel:';

beforeEach(async () => {
  failUpdate = false;
  localStorage.clear();
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

  // 注意：saveService 的 autosaver 固定 1500ms，这里用真实计时 + 略长于防抖周期的等待。
  it('notifyChanged 置脏，防抖到期仅写本地草稿：不动云端、不清脏', async () => {
    const docId = useEditorStore.getState().docId!;
    let latest: WorkbookSnapshot | null = null;
    bindEditor(docId, () => latest);

    notifyChanged();
    expect(useEditorStore.getState().dirty).toBe(true);

    latest = boundSnapshot();
    await new Promise((r) => setTimeout(r, 1800));

    // 云端内容未被自动保存触碰
    const cloud = await getDefaultStorage().load(docId);
    expect(cloud?.snapshot.sheets['sheet-01'].cellData[0]).toBeUndefined();
    // 本地草稿已写入最新快照，且 dirty 保持 true
    const draft = localStorage.getItem(DRAFT_PREFIX + docId);
    expect(draft).not.toBeNull();
    const parsed = JSON.parse(draft!) as WorkbookSnapshot;
    expect(parsed.sheets['sheet-01'].cellData[0]?.[0]?.v).toBe('当前内容');
    expect(useEditorStore.getState().dirty).toBe(true);
    expect(useEditorStore.getState().savedAt).toBeNull();
  }, 10_000);

  it('saveNow 云端保存成功：落库最新标题与快照、清 dirty 并删除草稿', async () => {
    const docId = useEditorStore.getState().docId!;
    let latest: WorkbookSnapshot | null = boundSnapshot();
    bindEditor(docId, () => latest);

    notifyChanged(); // dirty=true
    localStorage.setItem(DRAFT_PREFIX + docId, JSON.stringify(latest)); // 模拟防抖已写入草稿
    useEditorStore.getState().setTitle('手动保存的标题');
    await saveNow();

    const saved = await getDefaultStorage().load(docId);
    expect(saved?.title).toBe('手动保存的标题');
    expect(saved?.snapshot.sheets['sheet-01'].cellData[0]?.[0]?.v).toBe('当前内容');
    expect(useEditorStore.getState().dirty).toBe(false);
    expect(useEditorStore.getState().saving).toBe(false);
    expect(localStorage.getItem(DRAFT_PREFIX + docId)).toBeNull();
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

  it('写库失败时 saveNow 抛错且 saving 复位：dirty 保持、草稿保留兜底', async () => {
    failUpdate = true;
    const docId = useEditorStore.getState().docId!;
    let latest: WorkbookSnapshot | null = boundSnapshot();
    bindEditor(docId, () => latest);

    notifyChanged(); // dirty=true
    await expect(saveNow()).rejects.toThrow('mock 写库失败');
    expect(useEditorStore.getState().saving).toBe(false);
    expect(useEditorStore.getState().savedAt).toBeNull();
    expect(useEditorStore.getState().dirty).toBe(true);
    // 保存失败：⌘S 前写入的兜底草稿不被清除
    expect(localStorage.getItem(DRAFT_PREFIX + docId)).not.toBeNull();
  });

  it('unbindEditor 取消待执行的草稿写入', async () => {
    const docId = useEditorStore.getState().docId!;
    let latest: WorkbookSnapshot | null = boundSnapshot();
    bindEditor(docId, () => latest);

    notifyChanged();
    unbindEditor();
    await new Promise((r) => setTimeout(r, 1800));

    const saved = await getDefaultStorage().load(docId);
    expect(saved?.snapshot.sheets['sheet-01'].cellData[0]).toBeUndefined();
    expect(localStorage.getItem(DRAFT_PREFIX + docId)).toBeNull();
  }, 10_000);
});
