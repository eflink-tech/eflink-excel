import { indexedDbStorage } from './indexedDb';
import { EXCEL_DB_NAME } from './keys';
import type { StorageAdapter } from './types';

let adapter: StorageAdapter | null = null;

/** 注入全局默认存储（编辑器与文档操作都会使用它）。重复注入以最后一次为准。 */
export function setDefaultStorage(next: StorageAdapter): void {
  adapter = next;
}

/** 取当前默认存储；未注入时回退到内置 IndexedDB（库名 doculuma-excel，旧库数据自动搬运）。 */
export function getDefaultStorage(): StorageAdapter {
  adapter ??= indexedDbStorage(EXCEL_DB_NAME);
  return adapter;
}
