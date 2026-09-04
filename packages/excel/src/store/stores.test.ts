import { beforeEach, describe, expect, it } from 'vitest';
import { useEditorStore } from './editorStore';
import { useUiStore } from './uiStore';

function resetEditor() {
  useEditorStore.getState().openDoc({ id: 'doc-1', title: '初始' });
}

describe('editorStore 编辑器元状态', () => {
  beforeEach(resetEditor);

  it('openDoc 打开文档并清空保存状态', () => {
    useEditorStore.getState().markDirty();
    useEditorStore.getState().setSaving(true);
    useEditorStore.getState().markSaved();
    useEditorStore.getState().openDoc({ id: 'doc-2', title: '新文档' });

    const s = useEditorStore.getState();
    expect(s.docId).toBe('doc-2');
    expect(s.title).toBe('新文档');
    expect(s.dirty).toBe(false);
    expect(s.saving).toBe(false);
    expect(s.savedAt).toBeNull();
  });

  it('markDirty / setTitle / setSaving / markSaved 状态流转', () => {
    expect(useEditorStore.getState().dirty).toBe(false);
    useEditorStore.getState().markDirty();
    expect(useEditorStore.getState().dirty).toBe(true);

    useEditorStore.getState().setTitle('改名');
    expect(useEditorStore.getState().title).toBe('改名');

    useEditorStore.getState().setSaving(true);
    expect(useEditorStore.getState().saving).toBe(true);

    useEditorStore.getState().markSaved();
    const s = useEditorStore.getState();
    expect(s.dirty).toBe(false);
    expect(s.saving).toBe(false);
    expect(s.savedAt).not.toBeNull();
  });

  it('setDocId 单独切换文档 id 不影响标题', () => {
    useEditorStore.getState().setDocId('doc-9');
    expect(useEditorStore.getState().docId).toBe('doc-9');
    expect(useEditorStore.getState().title).toBe('初始');
  });
});

describe('uiStore 全局提示', () => {
  it('showToast / clearToast', () => {
    expect(useUiStore.getState().toast).toBeNull();
    useUiStore.getState().showToast('已保存');
    expect(useUiStore.getState().toast).toBe('已保存');
    useUiStore.getState().clearToast();
    expect(useUiStore.getState().toast).toBeNull();
  });
});
