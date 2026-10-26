/** DocuLuma 命名空间的本地存储键（唯一来源，业务代码不得再内联字符串）。 */

/** 本地草稿前缀：`doculuma:draft:excel:<docId>` */
export const DRAFT_PREFIX = 'doculuma:draft:excel:';

/** 图表元数据前缀：`doculuma:excel:chart-meta:<unitId>` */
export const CHART_META_PREFIX = 'doculuma:excel:chart-meta:';

/** 字体色 / 填充色「最近使用」的 localStorage key */
export const RECENT_COLORS_KEYS = {
  fontColor: 'doculuma:excel:recent-colors:fontColor',
  fillColor: 'doculuma:excel:recent-colors:fillColor',
} as const;

/** demo 宿主记录「上次文档」的 key（宿主侧使用，同时是迁移目标） */
export const LAST_DOC_KEY = 'doculuma:excel:last-doc';

/** IndexedDB 库名 */
export const EXCEL_DB_NAME = 'doculuma-excel';
export const EXCEL_AI_DB_NAME = 'doculuma-excel-ai';
