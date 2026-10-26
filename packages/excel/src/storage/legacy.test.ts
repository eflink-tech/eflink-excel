// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { migrateLegacyLocalStorage } from './legacy';
import { CHART_META_PREFIX, DRAFT_PREFIX, LAST_DOC_KEY, RECENT_COLORS_KEYS } from './keys';

beforeEach(() => {
  localStorage.clear();
});

describe('efLink → DocuLuma localStorage 命名空间迁移', () => {
  it('旧键值搬到新键，旧键被删除', () => {
    localStorage.setItem('eflink:draft:excel:doc-1', '{"a":1}');
    localStorage.setItem('eflink-excel:chart-meta:unit-1', '{"c":2}');
    localStorage.setItem('eflink-excel:recent-colors:fontColor', '["#fff"]');
    localStorage.setItem('eflink-excel:recent-colors:fillColor', '["#000"]');
    localStorage.setItem('eflink-excel:lastDoc', 'doc-1');

    migrateLegacyLocalStorage();

    expect(localStorage.getItem(DRAFT_PREFIX + 'doc-1')).toBe('{"a":1}');
    expect(localStorage.getItem(CHART_META_PREFIX + 'unit-1')).toBe('{"c":2}');
    expect(localStorage.getItem(RECENT_COLORS_KEYS.fontColor)).toBe('["#fff"]');
    expect(localStorage.getItem(RECENT_COLORS_KEYS.fillColor)).toBe('["#000"]');
    expect(localStorage.getItem(LAST_DOC_KEY)).toBe('doc-1');
    expect(localStorage.getItem('eflink:draft:excel:doc-1')).toBeNull();
    expect(localStorage.getItem('eflink-excel:chart-meta:unit-1')).toBeNull();
    expect(localStorage.getItem('eflink-excel:recent-colors:fontColor')).toBeNull();
    expect(localStorage.getItem('eflink-excel:recent-colors:fillColor')).toBeNull();
    expect(localStorage.getItem('eflink-excel:lastDoc')).toBeNull();
  });

  it('新键已有值时不覆盖，只清理旧键', async () => {
    // 全新模块实例：绕开「一个会话只迁移一次」的幂等守卫
    vi.resetModules();
    const { migrateLegacyLocalStorage: migrate } = await import('./legacy');

    localStorage.setItem('eflink:draft:excel:doc-2', '{"b":3}');
    localStorage.setItem(DRAFT_PREFIX + 'doc-2', '{"b":4}');

    expect(() => migrate()).not.toThrow();
    expect(localStorage.getItem(DRAFT_PREFIX + 'doc-2')).toBe('{"b":4}');
    expect(localStorage.getItem('eflink:draft:excel:doc-2')).toBeNull();
  });
});
