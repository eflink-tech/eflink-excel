import { normalizeLocale, type OfficeLocale } from '@eflink-tech/office-kit';
import { EXCEL_MESSAGES } from './messages';

/** React 组件之外的文案入口（文件动作、提示等），由 SheetEditor 同步当前语言 */
let currentLocale: OfficeLocale = 'zh-CN';

export function setExcelLocale(locale?: string): OfficeLocale {
  currentLocale = normalizeLocale(locale);
  return currentLocale;
}

export function getExcelLocale(): OfficeLocale {
  return currentLocale;
}

export function excelT(key: string, vars?: Record<string, string | number>): string {
  const template =
    EXCEL_MESSAGES[currentLocale]?.[key] ?? EXCEL_MESSAGES['zh-CN']?.[key] ?? key;
  if (!vars) return template;
  return template.replace(/\{(\w+)\}/g, (match, name: string) =>
    Object.prototype.hasOwnProperty.call(vars, name) ? String(vars[name]) : match,
  );
}
