/** 可本地化的业务错误：code 为 i18n 文案 key，message 仅用于日志/调试 */
export class OfficeError extends Error {
  readonly code: string;

  constructor(code: string, message?: string) {
    super(message ?? code);
    this.name = 'OfficeError';
    this.code = code;
  }
}

export function isOfficeError(error: unknown): error is OfficeError {
  return error instanceof OfficeError;
}

/** 把异常转成当前语言的提示；未知异常回退到 fallbackKey，避免泄露第三方英文/中文原文 */
export function resolveErrorMessage(
  error: unknown,
  t: (key: string) => string,
  fallbackKey: string,
): string {
  return isOfficeError(error) ? t(error.code) : t(fallbackKey);
}
