/** DocuLuma 命名空间迁移：把编辑器早期 efLink 前缀的本地数据一次性搬到 doculuma 前缀。 */

function getStorage(): Storage | null {
  try {
    const storage = globalThis.localStorage;
    if (!storage) return null;
    const probe = 'doculuma:storage-probe';
    storage.setItem(probe, '1');
    storage.removeItem(probe);
    return storage;
  } catch {
    // SSR 或隐私模式下 localStorage 不可用：迁移静默跳过，读写走新键即可
    return null;
  }
}

/** 单个键改名。新键已有值时只丢弃旧值，保证不覆盖本次会话写入的数据。 */
export function migrateLocalStorageKey(oldKey: string, newKey: string): void {
  const storage = getStorage();
  if (!storage) return;
  const value = storage.getItem(oldKey);
  if (value === null) return;
  if (storage.getItem(newKey) === null) storage.setItem(newKey, value);
  storage.removeItem(oldKey);
}

/** 前缀批量改名，用于 `eflink:draft:word:<id>` 这类按文档 id 展开的键。 */
export function migrateLocalStoragePrefix(oldPrefix: string, newPrefix: string): void {
  const storage = getStorage();
  if (!storage) return;
  const keys: string[] = [];
  for (let i = 0; i < storage.length; i += 1) {
    const key = storage.key(i);
    if (key?.startsWith(oldPrefix)) keys.push(key);
  }
  for (const key of keys) migrateLocalStorageKey(key, newPrefix + key.slice(oldPrefix.length));
}

/**
 * 读取旧库名的 IndexedDB 数据（只读，不删除旧库，失败返回 null）。
 * 各仓库拿到行数据后用自己的表类型写入新库，避免在这里耦合 Dexie 事务。
 */
export function readLegacyIndexedDb(
  oldName: string,
  storeNames: readonly string[],
): Promise<Record<string, unknown[]> | null> {
  if (typeof indexedDB === 'undefined') return Promise.resolve(null);
  return exists(oldName).then((had) => {
    if (!had) return null;
    return new Promise<Record<string, unknown[]> | null>((resolve) => {
      const request = indexedDB.open(oldName);
      request.onerror = () => resolve(null);
      request.onblocked = () => resolve(null);
      request.onupgradeneeded = () => {
        // 库不存在时 open() 会新建一个空库，立即回收这个占位符
        const db = request.result;
        db.close();
        indexedDB.deleteDatabase(oldName);
        resolve(null);
      };
      request.onsuccess = () => {
        const db = request.result;
        const stores = storeNames.filter((name) => db.objectStoreNames.contains(name));
        if (stores.length === 0) {
          db.close();
          resolve(null);
          return;
        }
        const tx = db.transaction(stores, 'readonly');
        const out: Record<string, unknown[]> = {};
        let pending = stores.length;
        for (const name of stores) {
          const all = tx.objectStore(name).getAll();
          all.onsuccess = () => {
            out[name] = all.result ?? [];
            if ((pending -= 1) === 0) {
              db.close();
              resolve(out);
            }
          };
          all.onerror = () => {
            db.close();
            resolve(null);
          };
        }
      };
    });
  });
}

/** `indexedDB.databases()` 并非所有浏览器都有，缺省时只能试探性 open()。 */
function exists(name: string): Promise<boolean> {
  const list = (indexedDB as unknown as { databases?: () => Promise<Array<{ name?: string }>> }).databases
    ?.();
  if (!list) return Promise.resolve(true);
  return list.then((all) => all.some((item) => item.name === name)).catch(() => true);
}
