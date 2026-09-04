import { create } from 'zustand';
import { getDefaultStorage } from '../storage/registry';
import { createDocument } from '../types/spreadsheet';
import type { DocumentMeta, SheetDocument } from '../types/spreadsheet';

interface DocumentsStore {
  docs: DocumentMeta[];
  refresh: () => Promise<void>;
  createDoc: (title?: string) => Promise<SheetDocument>;
  removeDoc: (id: string) => Promise<void>;
  renameDoc: (id: string, title: string) => Promise<void>;
  /** 导入等场景：外部已构造完整文档时落库 */
  addDoc: (doc: SheetDocument) => Promise<void>;
}

export const useDocumentsStore = create<DocumentsStore>((set, get) => ({
  docs: [],
  refresh: async () => set({ docs: await getDefaultStorage().list() }),
  createDoc: async (title) => {
    const doc = createDocument(title);
    await getDefaultStorage().save(doc);
    await get().refresh();
    return doc;
  },
  addDoc: async (doc) => {
    await getDefaultStorage().save(doc);
    await get().refresh();
  },
  removeDoc: async (id) => {
    await getDefaultStorage().delete(id);
    await get().refresh();
  },
  renameDoc: async (id, title) => {
    await getDefaultStorage().rename(id, title);
    await get().refresh();
  },
}));
