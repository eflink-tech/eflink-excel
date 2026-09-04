import { ChevronRight } from 'lucide-react';
import type { CSSProperties, ReactNode } from 'react';

/** 系统工具栏图标（myfsheet.css 中的 myf-icon-* / icon-*） */
export function ToolbarSysIcon({ iconClass }: { iconClass: string }) {
  return <span className={`myf-sys-item__icon-bg ${iconClass}`.trim()} aria-hidden />;
}

/** 系统风格下拉面板容器（圆角 + 轻阴影 + 分组间距） */
export function ToolbarDropdownPanel({
  children,
  className = '',
  style,
}: {
  children: ReactNode;
  className?: string;
  style?: CSSProperties;
}) {
  return (
    <div className={`myf-sys-panel ${className}`.trim()} style={style}>
      {children}
    </div>
  );
}

/** 系统风格菜单行：左图标 + 文案 + 可选快捷键/箭头 */
export function ToolbarMenuItem({
  label,
  icon,
  iconClass,
  right,
  chevron,
  active,
  onClick,
  onMouseEnter,
}: {
  label: string;
  icon?: ReactNode;
  iconClass?: string;
  right?: string;
  chevron?: boolean;
  active?: boolean;
  onClick?: () => void;
  onMouseEnter?: () => void;
}) {
  return (
    <button
      type="button"
      className={`myf-sys-item${active ? ' myf-sys-item--active' : ''}`}
      onClick={onClick}
      onMouseEnter={onMouseEnter}
    >
      <span className="myf-sys-item__icon">
        {icon ?? (iconClass ? <ToolbarSysIcon iconClass={iconClass} /> : null)}
      </span>
      <span className="myf-sys-item__label">{label}</span>
      {right != null && <span className="myf-sys-item__right">{right}</span>}
      {chevron && <ChevronRight size={12} className="myf-sys-item__chevron" aria-hidden />}
    </button>
  );
}

export function ToolbarMenuDivider() {
  return <div className="myf-sys-divider" role="separator" />;
}

/** 对齐类菜单行：选中勾 + 对齐图标 + 文案 */
export function ToolbarAlignMenuItem({
  label,
  selected,
  alignIcon,
  onClick,
}: {
  label: string;
  selected: boolean;
  alignIcon: string;
  onClick: () => void;
}) {
  return (
    <button type="button" className="myf-sys-item myf-sys-item--align" onClick={onClick}>
      <span className={`myf-sys-item__icon-bg myf-x-ico${selected ? ' myf-icon-selected' : ''}`.trim()} aria-hidden />
      <span className={`myf-sys-item__icon-bg ${alignIcon}`.trim()} aria-hidden />
      <span className="myf-sys-item__label">{label}</span>
    </button>
  );
}
