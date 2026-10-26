/** 五个编辑器共用的界面语言集合。默认 zh-CN，保持 eflink.tech 现有观感不变。 */
export const OFFICE_LOCALES = ['zh-CN', 'en', 'es', 'pt-BR'] as const;

export type OfficeLocale = (typeof OFFICE_LOCALES)[number];

export const DEFAULT_LOCALE: OfficeLocale = 'zh-CN';

export const LOCALE_LABELS: Record<OfficeLocale, string> = {
  'zh-CN': '简体中文',
  en: 'English',
  es: 'Español',
  'pt-BR': 'Português (Brasil)',
};

/** 兼容 en-US / es-MX / pt-br / zh 等宿主传入的宽松写法 */
export function normalizeLocale(input?: string | null): OfficeLocale {
  const value = (input ?? '').trim().toLowerCase();
  if (!value) return DEFAULT_LOCALE;
  if (value.startsWith('zh')) return 'zh-CN';
  if (value.startsWith('es')) return 'es';
  if (value.startsWith('pt')) return 'pt-BR';
  if (value.startsWith('en')) return 'en';
  return (OFFICE_LOCALES as readonly string[]).includes(value)
    ? (value as OfficeLocale)
    : DEFAULT_LOCALE;
}
