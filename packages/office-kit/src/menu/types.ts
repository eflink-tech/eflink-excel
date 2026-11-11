import type { ReactNode } from 'react';

export interface OfficeMenuItem {
  key: string;
  label: string;
  /** 直接传图节点（如 <FileText size={16} />），office-kit 不绑定图标库 */
  icon?: ReactNode;
  shortcut?: string;
  checked?: boolean;
  disabled?: boolean;
  danger?: boolean;
  onSelect?: () => void;
  /** 子菜单按分组组织，组间显示分隔线 */
  children?: OfficeMenuItem[][];
}

export interface OfficeMenu {
  key: string;
  label: string;
  items: OfficeMenuItem[][];
  /** 下拉对齐方式；默认左对齐（下拉从按钮左缘向右展开） */
  align?: 'start' | 'end';
}
