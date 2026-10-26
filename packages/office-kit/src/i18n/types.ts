import type { OfficeLocale } from './locales';

export type OfficeMessages = Record<string, string>;

export type OfficeMessagesByLocale = Partial<Record<OfficeLocale, OfficeMessages>>;

export type TranslateVars = Record<string, string | number>;

export interface I18nValue {
  locale: OfficeLocale;
  t: (key: string, vars?: TranslateVars) => string;
}
