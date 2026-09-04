import { ChevronRight, Keyboard, Pencil, Save } from 'lucide-react';
import { cloneElement, isValidElement, useEffect, useRef, useState, type ReactElement, type ReactNode } from 'react';
import { ShortcutsDialog } from './ShortcutsDialog';
import { ToolbarDropdownPanel, ToolbarMenuDivider, ToolbarMenuItem, ToolbarSysIcon } from './ToolbarMenuPanel';
import './myfsheet.css';
import { closeAllToolbarDropdowns, registerToolbarDropdown } from './toolbarDropdown';
import {
  exportEfexcelAction,
  exportPngAction,
  newDocAction,
  saveAction,
} from '../../core/fileActions';
import * as cmd from '../../core/univer/commands';

interface MenuLeaf {
  label: string;
  icon?: ReactNode;
  iconClass?: string;
  shortcut?: string;
  action: () => void;
}

interface MenuBranch {
  label: string;
  icon?: ReactNode;
  iconClass?: string;
  children: MenuLeaf[];
}

type MenuItem = MenuBranch | MenuLeaf;
function isBranch(item: MenuItem): item is MenuBranch {
  return 'children' in item;
}

function useMenuItems(importRef: React.RefObject<HTMLInputElement | null>): MenuItem[] {
  const fileChildren: MenuLeaf[] = [
    { label: '新建表格', iconClass: 'myf-icon-add', action: () => void newDocAction() },
    { label: '保存', icon: <Save size={16} />, shortcut: '⌘S', action: () => void saveAction() },
    { label: '导入表格(.efexcel)', iconClass: 'myf-icon-file-import', action: () => importRef.current?.click() },
    { label: '导出表格(.efexcel)', iconClass: 'myf-icon-file-export', action: () => void exportEfexcelAction() },
    { label: '导出图片', iconClass: 'myf-icon-file-download', action: () => void exportPngAction() },
  ];
  const editChildren: MenuLeaf[] = [
    { label: '撤销', iconClass: 'myf-icon-undo', action: cmd.undo },
    { label: '重做', iconClass: 'myf-icon-redo', action: cmd.redo },
    { label: '清除内容', iconClass: 'myf-icon-clear-format', action: cmd.clearContentOnly },
    { label: '清除格式', iconClass: 'myf-icon-clear-format', action: cmd.clearFormatOnly },
  ];
  const viewChildren: MenuLeaf[] = [
    { label: '冻结首行', iconClass: 'myf-icon-freeze', action: cmd.freezeFirstRow },
    { label: '冻结首列', iconClass: 'myf-icon-freeze', action: cmd.freezeFirstColumn },
    { label: '取消冻结', iconClass: 'myf-icon-freeze', action: cmd.cancelFreeze },
  ];

  return [
    { label: '文件', iconClass: 'myf-icon-file-menu', children: fileChildren },
    { label: '编辑', icon: <Pencil size={16} />, children: editChildren },
    { label: '查看', iconClass: 'myf-icon-freeze', children: viewChildren },
    {
      label: '快捷键',
      icon: <Keyboard size={16} />,
      action: () => { /* handled externally */ },
    },
  ];
}

function mergeTriggerClick(trigger: ReactNode, onToggle: () => void): ReactNode {
  if (!isValidElement(trigger)) return trigger;
  const el = trigger as ReactElement<{ onClick?: (e: React.MouseEvent) => void }>;
  return cloneElement(el, {
    onClick: (e: React.MouseEvent) => {
      el.props.onClick?.(e);
      if (!e.defaultPrevented) onToggle();
    },
  });
}

/** 触发按钮 + 二级下拉菜单（TitleBar / SheetToolbar 共用） */
export function MenuDropdown({
  trigger,
  importInputRef,
  variant = 'default',
  dropdownId = 'main-menu',
  onOpen,
}: {
  trigger: ReactNode;
  importInputRef: React.RefObject<HTMLInputElement | null>;
  variant?: 'default' | 'sheet';
  dropdownId?: string;
  onOpen?: () => void;
}) {
  const [open, setOpen] = useState(false);
  const [active, setActive] = useState<MenuItem | null>(null);
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

  const items = useMenuItems(importInputRef);

  useEffect(() => {
    return registerToolbarDropdown(dropdownId, () => {
      setOpen(false);
      setActive(null);
    });
  }, [dropdownId]);

  useEffect(() => {
    if (!open) return;
    const onClick = (e: MouseEvent) => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) {
        setOpen(false);
        setActive(null);
      }
    };
    document.addEventListener('click', onClick, true);
    return () => document.removeEventListener('click', onClick, true);
  }, [open]);

  const toggle = () => {
    setOpen((v) => {
      const next = !v;
      if (next) {
        closeAllToolbarDropdowns(dropdownId);
        onOpen?.();
      }
      setActive(null);
      return next;
    });
  };

  const run = (fn: () => void) => {
    setOpen(false);
    setActive(null);
    fn();
  };

  const triggerNode = mergeTriggerClick(trigger, toggle);

  const renderMenuIcon = (item: { icon?: ReactNode; iconClass?: string }) => {
    if (item.iconClass) return <ToolbarSysIcon iconClass={item.iconClass} />;
    if (item.icon) return <span className="text-slate-500">{item.icon}</span>;
    return null;
  };

  const renderSheetMenu = () => (
    <div className="myf-sys-menu-row">
      <ToolbarDropdownPanel className="myf-sys-panel--static" style={{ width: 148, position: 'static', marginTop: 0 }}>
        {items.map((item, i) => (
          <div key={item.label}>
            {i === items.length - 1 && <ToolbarMenuDivider />}
            <ToolbarMenuItem
              label={item.label}
              icon={item.iconClass ? undefined : item.icon}
              iconClass={item.iconClass}
              chevron={isBranch(item)}
              active={isBranch(item) && active === item}
              onMouseEnter={() => setActive(isBranch(item) ? item : null)}
              onClick={() => {
                if (isBranch(item)) setActive(item);
                else if (item.label === '快捷键') {
                  setOpen(false);
                  setActive(null);
                  setShortcutsOpen(true);
                } else run(item.action);
              }}
            />
          </div>
        ))}
      </ToolbarDropdownPanel>
      {active && isBranch(active) && (
        <ToolbarDropdownPanel className="myf-sys-submenu myf-sys-panel--static" style={{ width: 224, position: 'static', marginTop: 0 }}>
          {active.children.map((leaf) => (
            <div key={leaf.label}>
              {leaf.label === '清除内容' && <ToolbarMenuDivider />}
              <ToolbarMenuItem
                label={leaf.label}
                icon={leaf.iconClass ? undefined : leaf.icon}
                iconClass={leaf.iconClass}
                right={leaf.shortcut}
                onClick={() => run(leaf.action)}
              />
            </div>
          ))}
        </ToolbarDropdownPanel>
      )}
    </div>
  );

  if (variant === 'sheet') {
    return (
      <div ref={rootRef} className="relative">
        {triggerNode}
        {open && renderSheetMenu()}
        <ShortcutsDialog open={shortcutsOpen} onClose={() => setShortcutsOpen(false)} />
      </div>
    );
  }

  return (
    <div ref={rootRef} className="relative">
      {triggerNode}

      {open && (
        <div className="absolute left-0 top-full z-[200] mt-2 flex items-start">
          <div className="w-52 rounded-lg border border-slate-200 bg-white py-1 shadow-xl">
            {items.map((item, i) => {
              const isLast = i === items.length - 1;
              return (
                <button
                  key={item.label}
                  type="button"
                  className={`flex w-full items-center gap-2.5 px-3 py-2 text-left text-[13px] text-slate-700 transition-colors hover:bg-[#f2f3f4] ${
                    isLast ? 'mt-1 border-t border-slate-100 pt-2' : ''
                  } ${isBranch(item) && active === item ? 'bg-[#f2f3f4]' : ''}`}
                  onMouseEnter={() => setActive(isBranch(item) ? item : null)}
                  onClick={() => {
                    if (isBranch(item)) setActive(item);
                    else if (item.label === '快捷键') {
                      setOpen(false);
                      setActive(null);
                      setShortcutsOpen(true);
                    } else run(item.action);
                  }}
                >
                  {renderMenuIcon(item)}
                  <span className="flex-1">{item.label}</span>
                  {isBranch(item) && <ChevronRight size={14} className="text-slate-400" />}
                </button>
              );
            })}
          </div>

          {active && isBranch(active) && (
            <div className="ml-1 w-56 rounded-lg border border-slate-200 bg-white py-1 shadow-xl">
              {active.children.map((leaf) => (
                <button
                  key={leaf.label}
                  type="button"
                  className="flex w-full items-center gap-2.5 px-3 py-[7px] text-left text-[13px] text-slate-700 transition-colors hover:bg-[#f2f3f4]"
                  onClick={() => run(leaf.action)}
                >
                  {renderMenuIcon(leaf)}
                  <span className="flex-1">{leaf.label}</span>
                  {leaf.shortcut && <span className="text-xs text-slate-400">{leaf.shortcut}</span>}
                </button>
              ))}
            </div>
          )}
        </div>
      )}

      <ShortcutsDialog open={shortcutsOpen} onClose={() => setShortcutsOpen(false)} />
    </div>
  );
}
