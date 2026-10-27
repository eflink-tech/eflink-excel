import Dexie from 'dexie';
import { isLegacyImportDone, markLegacyImportDone, readLegacyIndexedDb } from '@eflink-tech/office-kit';
import type { DocumentMeta, SheetDocument, WorkbookSnapshot } from '../types/spreadsheet';
import { EXCEL_DB_NAME } from './keys';
import type { StorageAdapter } from './types';

/** Dexie package `exports` omit types; cast keeps IndexedDB API typed under bundler resolution. */
type DocumentsTable = {
  put(doc: SheetDocument): Promise<string>;
  get(id: string): Promise<SheetDocument | undefined>;
  delete(id: string): Promise<void>;
  clear(): Promise<void>;
  count(): Promise<number>;
  bulkPut(docs: SheetDocument[]): Promise<unknown>;
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

/** 旧库名（efLink 品牌时期）：只读搬运，绝不删除，用户可随时回退旧版本 */
const LEGACY_DB_NAME = 'eflink-excel';
const LEGACY_DOCUMENTS_STORE = 'documents';

const legacyImportCache = new Map<string, Promise<void>>();

/**
 * 一次性把旧库文档搬进新库：仅在目标表为空时写入，失败只告警、不阻断编辑器。
 * 按目标库名 memoize，同一会话只跑一次。
 */
export function ensureLegacyImport(db: SheetDB, dbName: string): Promise<void> {
  let pending = legacyImportCache.get(dbName);
  if (!pending) {
    pending = importLegacyDocuments(db, dbName).catch((err) => {
      console.warn('[indexedDb] 旧库数据搬运失败', err);
    });
    legacyImportCache.set(dbName, pending);
  }
  return pending;
}

async function importLegacyDocuments(db: SheetDB, dbName: string): Promise<void> {
  // 宿主显式沿用旧库名时无需搬运（数据本就在目标库）
  if (dbName === LEGACY_DB_NAME) return;
  // 一次性标记防「删光复活」：空表判据单独使用时，用户删光文档后每次启动都会回填旧数据
  const markerKey = `doculuma:excel:legacy-imported:${dbName}:documents`;
  if (isLegacyImportDone(markerKey)) return;
  if ((await db.documents.count()) > 0) {
    // 新库已有数据：以新库为准，落标记使「跳过」也成为一次性语义
    markLegacyImportDone(markerKey);
    return;
  }
  const legacy = await readLegacyIndexedDb(LEGACY_DB_NAME, [LEGACY_DOCUMENTS_STORE]);
  const rows = legacy?.[LEGACY_DOCUMENTS_STORE] as SheetDocument[] | undefined;
  if (!rows?.length) {
    markLegacyImportDone(markerKey);
    return;
  }
  await db.documents.bulkPut(rows);
  markLegacyImportDone(markerKey);
}

/** 内置 IndexedDB 存储（Dexie）。dbName 缺省为 doculuma-excel，旧 eflink-excel 数据会自动搬运。 */
export function indexedDbStorage(dbName = EXCEL_DB_NAME): StorageAdapter {
  const db = new SheetDB(dbName);
  return {
    async save(doc) {
      await db.documents.put(doc);
    },
    async load(id) {
      await ensureLegacyImport(db, dbName);
      return db.documents.get(id);
    },
    async list(): Promise<DocumentMeta[]> {
      await ensureLegacyImport(db, dbName);
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
