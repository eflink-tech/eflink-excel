import { create } from 'zustand';

// 编辑器元状态：快照本体由 Univer 实例持有（saveService 持有取快照的引用），store 只管标题与保存状态
interface EditorStore {
  docId: string | null;
  title: string;
  dirty: boolean;
  saving: boolean;
  savedAt: number | null;
  openDoc: (doc: { id: string; title: string }) => void;
  setDocId: (id: string | null) => void;
  markDirty: () => void;
  setTitle: (title: string) => void;
  setSaving: (saving: boolean) => void;
  markSaved: () => void;
}

export const useEditorStore = create<EditorStore>((set) => ({
  docId: null,
  title: '',
  dirty: false,
  saving: false,
  savedAt: null,
  openDoc: ({ id, title }) => set({ docId: id, title, dirty: false, saving: false, savedAt: null }),
  setDocId: (id) => set({ docId: id }),
  markDirty: () => set({ dirty: true }),
  setTitle: (title) => set({ title }),
  setSaving: (saving) => set({ saving }),
  markSaved: () => set({ dirty: false, saving: false, savedAt: Date.now() }),
}));
