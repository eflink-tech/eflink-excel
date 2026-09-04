import Dexie from 'dexie';
import type { DocumentMeta, SheetDocument, WorkbookSnapshot } from '../types/spreadsheet';
import type { StorageAdapter } from './types';

/** Dexie package `exports` omit types; cast keeps IndexedDB API typed under bundler resolution. */
type DocumentsTable = {
  put(doc: SheetDocument): Promise<string>;
  get(id: string): Promise<SheetDocument | undefined>;
  delete(id: string): Promise<void>;
  clear(): Promise<void>;
  update(id: string, changes: Partial<SheetDocument>): Promise<number>;
  orderBy(index: keyof SheetDocument | string): {
    reverse(): { toArray(): Promise<SheetDocument[]> };
  };
};

type SheetDBInstance = {
  version(v: number): { stores(schema: Record<string, string | null>): unknown };
  documents: DocumentsTable;
};

const DexieBase = Dexie as unknown as new (name: string) => SheetDBInstance;

class SheetDB extends DexieBase {
  declare documents: DocumentsTable;

  constructor(name: string) {
    super(name);
    this.version(1).stores({ documents: 'id, title, updatedAt' });
  }
}

/** 内置 IndexedDB 存储（Dexie）。dbName 缺省与历史单体版一致，老数据可直接读回。 */
export function indexedDbStorage(dbName = 'eflink-excel'): StorageAdapter {
  const db = new SheetDB(dbName);
  return {
    async save(doc) {
      await db.documents.put(doc);
    },
    async load(id) {
      return db.documents.get(id);
    },
    async list(): Promise<DocumentMeta[]> {
      const all = await db.documents.orderBy('updatedAt').reverse().toArray();
      return all.map((doc) => ({ id: doc.id, title: doc.title, updatedAt: doc.updatedAt }));
    },
    async delete(id) {
      await db.documents.delete(id);
    },
    async rename(id, title) {
      await db.documents.update(id, { title, updatedAt: Date.now() });
    },
    async updateContent(id, title, snapshot: WorkbookSnapshot) {
      await db.documents.update(id, { title, snapshot, updatedAt: Date.now() });
    },
    async clear() {
      await db.documents.clear();
    },
  };
}
