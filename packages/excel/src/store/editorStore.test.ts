import { describe, expect, it } from 'vitest';
import { useEditorStore } from './editorStore';

describe('editorStore reloadToken', () => {
  it('bumpReload 递增重载计数', () => {
    const before = useEditorStore.getState().reloadToken;
    useEditorStore.getState().bumpReload();
    expect(useEditorStore.getState().reloadToken).toBe(before + 1);
  });
});
