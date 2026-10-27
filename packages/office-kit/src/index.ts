import './styles.css';

export {
  OFFICE_LOCALES,
  DEFAULT_LOCALE,
  LOCALE_LABELS,
  normalizeLocale,
  type OfficeLocale,
} from './i18n/locales';
export type { OfficeMessages, OfficeMessagesByLocale, TranslateVars, I18nValue } from './i18n/types';
export { COMMON_MESSAGES } from './i18n/dictionaries';
export { I18nProvider, useI18n, useTranslate, registerOfficeMessages } from './i18n/I18nProvider';
export type { I18nProviderProps } from './i18n/I18nProvider';

export { OfficeError, isOfficeError, resolveErrorMessage } from './errors/OfficeError';

export { OfficeMenuBar } from './menu/OfficeMenuBar';
export type { OfficeMenuBarProps } from './menu/OfficeMenuBar';
export type { OfficeMenu, OfficeMenuItem } from './menu/types';

export { OfficeHeader } from './header/OfficeHeader';
export type { OfficeHeaderBrand, OfficeHeaderProps } from './header/OfficeHeader';

export {
  migrateLocalStorageKey,
  migrateLocalStoragePrefix,
  readLegacyIndexedDb,
  isLegacyImportDone,
  markLegacyImportDone,
} from './storage/migrate';
