// 品牌更名（efLink 易飞 → DocuLuma）后的一次性本地数据迁移。
// 这里保留的旧 efLink 字符串只服务于「读取旧键」，业务代码一律用 ./keys 里的新常量。
import { migrateLocalStorageKey, migrateLocalStoragePrefix } from '@eflink-tech/office-kit';
import {
  CHART_META_PREFIX,
  DRAFT_PREFIX,
  LAST_DOC_KEY,
  RECENT_COLORS_KEYS,
} from './keys';

/** 旧 localStorage 键/前缀（efLink 命名空间） */
const LEGACY_DRAFT_PREFIX = 'eflink:draft:excel:';
const LEGACY_CHART_META_PREFIX = 'eflink-excel:chart-meta:';
const LEGACY_RECENT_FONT_COLOR_KEY = 'eflink-excel:recent-colors:fontColor';
const LEGACY_RECENT_FILL_COLOR_KEY = 'eflink-excel:recent-colors:fillColor';
const LEGACY_LAST_DOC_KEY = 'eflink-excel:lastDoc';

let hasMigrated = false;

/**
 * 把 efLink 命名空间的 localStorage 数据一次性改名到 doculuma 命名空间。
 * 真改名（新键已存在则只丢弃旧值）、幂等、localStorage 不可用时静默跳过，绝不抛错。
 * 由 src/index.ts 在模块加载时调用一次。
 */
export function migrateLegacyLocalStorage(): void {
  if (hasMigrated) return;
  hasMigrated = true;
  try {
    migrateLocalStoragePrefix(LEGACY_DRAFT_PREFIX, DRAFT_PREFIX);
    migrateLocalStoragePrefix(LEGACY_CHART_META_PREFIX, CHART_META_PREFIX);
    migrateLocalStorageKey(LEGACY_RECENT_FONT_COLOR_KEY, RECENT_COLORS_KEYS.fontColor);
    migrateLocalStorageKey(LEGACY_RECENT_FILL_COLOR_KEY, RECENT_COLORS_KEYS.fillColor);
    migrateLocalStorageKey(LEGACY_LAST_DOC_KEY, LAST_DOC_KEY);
  } catch (err) {
    // 迁移只是兜底，失败也不能阻断编辑器启动
    console.warn('[legacy] localStorage 命名空间迁移失败', err);
  }
}
