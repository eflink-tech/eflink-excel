import { createContext, useCallback, useContext, useMemo } from 'react';
import type { ReactNode } from 'react';
import { COMMON_MESSAGES } from './dictionaries';
import { DEFAULT_LOCALE, normalizeLocale, type OfficeLocale } from './locales';
import type { I18nValue, OfficeMessagesByLocale, TranslateVars } from './types';

const FALLBACK_LOCALE: OfficeLocale = DEFAULT_LOCALE;

function interpolate(template: string, vars?: TranslateVars): string {
  if (!vars) return template;
  return template.replace(/\{(\w+)\}/g, (match, name: string) =>
    Object.prototype.hasOwnProperty.call(vars, name) ? String(vars[name]) : match,
  );
}

/** 编辑器包在模块加载时注册自己的默认文案，供脱离 Provider 渲染的组件兜底 */
let registeredMessages: OfficeMessagesByLocale | null = null;

export function registerOfficeMessages(messages: OfficeMessagesByLocale): void {
  registeredMessages = messages;
}

function buildFallbackValue(): I18nValue {
  const dict = {
    ...COMMON_MESSAGES[FALLBACK_LOCALE],
    ...registeredMessages?.[FALLBACK_LOCALE],
  };
  return {
    locale: FALLBACK_LOCALE,
    t: (key, vars) => interpolate(dict[key] ?? key, vars),
  };
}

const I18nContext = createContext<I18nValue | null>(null);

export interface I18nProviderProps {
  /** 宽松接受 en-US / es-MX / pt-br 等写法；缺省 zh-CN */
  locale?: string;
  /** 编辑器专属文案；按 key 覆盖共用文案 */
  messages?: OfficeMessagesByLocale;
  children: ReactNode;
}

export function I18nProvider({ locale, messages, children }: I18nProviderProps) {
  const resolved = normalizeLocale(locale);

  const value = useMemo<I18nValue>(() => {
    const dict = {
      ...COMMON_MESSAGES[resolved],
      ...messages?.[resolved],
    };
    const fallback = {
      ...COMMON_MESSAGES[FALLBACK_LOCALE],
      ...messages?.[FALLBACK_LOCALE],
    };
    const t = (key: string, vars?: TranslateVars) =>
      interpolate(dict[key] ?? fallback[key] ?? key, vars);
    return { locale: resolved, t };
  }, [resolved, messages]);

  return <I18nContext.Provider value={value}>{children}</I18nContext.Provider>;
}

export function useI18n(): I18nValue {
  const context = useContext(I18nContext);
  const fallback = useMemo(buildFallbackValue, []);
  return context ?? fallback;
}

/** 跨越组件边界传递稳定的 t（用于 useCallback 依赖） */
export function useTranslate(): I18nValue['t'] {
  const { t } = useI18n();
  return useCallback((key: string, vars?: TranslateVars) => t(key, vars), [t]);
}
