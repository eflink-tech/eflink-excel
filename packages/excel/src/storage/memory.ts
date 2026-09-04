import type { StorageAdapter } from './types';

/** 内存存储：无持久化（刷新即丢），适合嵌入式无痕场景与单元测试。 */
export function memoryStorage(): StorageAdapter {
  const docs = new Map<string, string>(); // 借 JSON 序列化做深拷贝，避免引用被外部改写
  return {
    async save(doc) {
      docs.set(doc.id, JSON.stringify(doc));
    },
    async load(id) {
      const raw = docs.get(id);
      return raw ? (JSON.parse(raw) as Awaited<ReturnType<StorageAdapter['load']>>) : undefined;
    },
    async list() {
      const all = await Promise.all([...docs.keys()].map((id) => this.load(id)));
      return all
        .filter((doc): doc is NonNullable<typeof doc> => Boolean(doc))
        .map(({ id, title, updatedAt }) => ({ id, title, updatedAt }))
        .sort((a, b) => b.updatedAt - a.updatedAt);
    },
    async delete(id) {
      docs.delete(id);
    },
    async rename(id, title) {
      const doc = await this.load(id);
      if (doc) await this.save({ ...doc, title, updatedAt: Date.now() });
    },
    async updateContent(id, title, snapshot) {
      const doc = await this.load(id);
      if (doc) await this.save({ ...doc, title, snapshot, updatedAt: Date.now() });
    },
    async clear() {
      docs.clear();
    },
  };
}
