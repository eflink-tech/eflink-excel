/** 工具栏下拉互斥：打开一个时关闭其他已注册的下拉 */
type CloseFn = () => void;
const registry = new Map<string, CloseFn>();

export function registerToolbarDropdown(id: string, close: CloseFn): () => void {
  registry.set(id, close);
  return () => registry.delete(id);
}

export function closeAllToolbarDropdowns(exceptId?: string): void {
  for (const [id, close] of registry) {
    if (id !== exceptId) close();
  }
}
