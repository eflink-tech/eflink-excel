// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';
import { clearDraft, readDraft, writeDraft } from './draft';
import { createEmptySnapshot } from '../types/spreadsheet';

beforeEach(() => {
  localStorage.clear();
});

describe('本地草稿（localStorage 兜底）', () => {
  it('writeDraft 按 eflink:draft:excel:<docId> 键写入序列化快照', () => {
    const snapshot = createEmptySnapshot();
    writeDraft('doc-1', snapshot);
    const raw = localStorage.getItem('eflink:draft:excel:doc-1');
    expect(raw).not.toBeNull();
    expect(JSON.parse(raw!)).toEqual(snapshot);
  });

  it('readDraft 读回写入的快照，无草稿时返回 null', () => {
    expect(readDraft('doc-1')).toBeNull();
    const snapshot = createEmptySnapshot();
    writeDraft('doc-1', snapshot);
    expect(readDraft('doc-1')).toEqual(snapshot);
  });

  it('clearDraft 删除草稿', () => {
    writeDraft('doc-1', createEmptySnapshot());
    clearDraft('doc-1');
    expect(readDraft('doc-1')).toBeNull();
  });

  it('草稿损坏时 readDraft 返回 null 而非抛错', () => {
    localStorage.setItem('eflink:draft:excel:bad', '{broken json');
    expect(readDraft('bad')).toBeNull();
  });
});
