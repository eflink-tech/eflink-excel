import type { ReactNode } from 'react';
import { useI18n } from '../i18n/I18nProvider';
import { OfficeMenuBar } from '../menu/OfficeMenuBar';
import type { OfficeMenu } from '../menu/types';

export interface OfficeHeaderBrand {
  logo?: string;
  name: string;
}

export interface OfficeHeaderProps {
  brand?: OfficeHeaderBrand | false;
  backHref?: string | null;
  menus?: OfficeMenu[];
  actions?: ReactNode;
  /** 便于宿主定位/覆盖样式 */
  className?: string;
}

function BackIcon() {
  return (
    <svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true">
      <path
        d="M11.5 4.5 6 10l5.5 5.5"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      <path
        d="M6.5 10H15"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
      />
    </svg>
  );
}

/** 统一顶栏：返回 + 品牌 + 横向菜单 + 右侧动作（分享/反馈/AI 等） */
export function OfficeHeader({ brand, backHref, menus = [], actions, className }: OfficeHeaderProps) {
  const { t } = useI18n();
  return (
    <header className={`efk-header${className ? ` ${className}` : ''}`} role="banner">
      {(backHref || brand) && (
        <div className="efk-brand">
          {backHref && (
            <a
              href={backHref}
              className="efk-back"
              title={t('shell.back')}
              aria-label={t('shell.back')}
            >
              <BackIcon />
            </a>
          )}
          {brand && brand.logo && <img className="efk-brand-logo" src={brand.logo} alt="" />}
          {brand && <span className="efk-brand-name">{brand.name}</span>}
        </div>
      )}
      {menus.length > 0 && <OfficeMenuBar menus={menus} className="efk-header-menus" />}
      {actions && <div className="efk-header-actions">{actions}</div>}
    </header>
  );
}
