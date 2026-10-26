import {
  Fragment,
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import type { KeyboardEvent as ReactKeyboardEvent } from 'react';
import { useI18n } from '../i18n/I18nProvider';
import type { OfficeMenu, OfficeMenuItem } from './types';

const GAP = 2;
const MORE_KEY = '__more';

function estimateWidth(label: string): number {
  return Math.max(40, Math.round(label.length * 9 + 20));
}

function CheckMark() {
  return (
    <svg viewBox="0 0 16 16" width="13" height="13" aria-hidden="true" className="efk-check">
      <path
        d="M3.5 8.5 6.5 11.5 12.5 5"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function ChevronRight() {
  return (
    <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true" className="efk-chevron">
      <path
        d="M6 3.5 10.5 8 6 12.5"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function MoreDots() {
  return (
    <svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true">
      <circle cx="4" cy="10" r="1.6" fill="currentColor" />
      <circle cx="10" cy="10" r="1.6" fill="currentColor" />
      <circle cx="16" cy="10" r="1.6" fill="currentColor" />
    </svg>
  );
}

interface MenuRowsProps {
  groups: OfficeMenuItem[][];
  openSub: string | null;
  setOpenSub: (key: string | null) => void;
  onSelect: (item: OfficeMenuItem) => void;
  flipSubmenu: boolean;
}

function MenuRows({ groups, openSub, setOpenSub, onSelect, flipSubmenu }: MenuRowsProps) {
  return (
    <>
      {groups.map((group, groupIndex) => (
        <Fragment key={groupIndex}>
          {groupIndex > 0 && <div className="efk-menu-divider" role="separator" />}
          {group.map((item) => {
            const hasChildren = Boolean(item.children?.length);
            return (
              <div
                key={item.key}
                className="efk-menu-row"
                onMouseEnter={() => setOpenSub(hasChildren ? item.key : null)}
              >
                <button
                  type="button"
                  role={item.checked === undefined ? 'menuitem' : 'menuitemcheckbox'}
                  aria-checked={item.checked}
                  aria-haspopup={hasChildren ? 'menu' : undefined}
                  aria-expanded={hasChildren ? openSub === item.key : undefined}
                  disabled={item.disabled}
                  className={`efk-menu-item${item.checked ? ' is-checked' : ''}${
                    item.danger ? ' is-danger' : ''
                  }`}
                  onClick={() => {
                    if (hasChildren) {
                      setOpenSub(openSub === item.key ? null : item.key);
                      return;
                    }
                    onSelect(item);
                  }}
                >
                  <span className="efk-menu-icon">{item.icon}</span>
                  <span className="efk-menu-label">{item.label}</span>
                  {item.checked && <CheckMark />}
                  {item.shortcut && <span className="efk-menu-shortcut">{item.shortcut}</span>}
                  {hasChildren && <ChevronRight />}
                </button>
                {hasChildren && openSub === item.key && (
                  <div
                    className={`efk-submenu${flipSubmenu ? ' efk-submenu--left' : ''}`}
                    role="menu"
                  >
                    <MenuRows
                      groups={item.children!}
                      openSub={openSub}
                      setOpenSub={setOpenSub}
                      onSelect={onSelect}
                      flipSubmenu={flipSubmenu}
                    />
                  </div>
                )}
              </div>
            );
          })}
        </Fragment>
      ))}
    </>
  );
}

export interface OfficeMenuBarProps {
  menus: OfficeMenu[];
  className?: string;
  ariaLabel?: string;
}

/**
 * 横向菜单栏：点击展开、已展开时悬停切换；宽度不足时靠右折叠进「更多」。
 * 子菜单悬停展开，点击叶子项后收起整组菜单。
 */
export function OfficeMenuBar({ menus, className, ariaLabel }: OfficeMenuBarProps) {
  const { t } = useI18n();
  const barRef = useRef<HTMLDivElement>(null);
  const [openKey, setOpenKey] = useState<string | null>(null);
  const [openSub, setOpenSub] = useState<string | null>(null);
  const [barWidth, setBarWidth] = useState(0);
  const [, setMeasureVersion] = useState(0);
  const widthCache = useRef<Record<string, number>>({});
  // 菜单数组通常每次渲染重建，用签名替代身份依赖，避免反复创建 ResizeObserver
  const menuSignature = menus.map((menu) => `${menu.key}:${menu.label}`).join('|');

  const measure = useCallback(() => {
    const root = barRef.current;
    if (!root) return;
    let changed = false;
    root.querySelectorAll<HTMLElement>('[data-efk-menu-key]').forEach((node) => {
      const key = node.dataset.efkMenuKey;
      const width = node.offsetWidth;
      if (key && width > 0 && widthCache.current[key] !== width) {
        widthCache.current[key] = width;
        changed = true;
      }
    });
    if (changed) setMeasureVersion((value) => value + 1);
  }, []);

  useLayoutEffect(() => {
    const root = barRef.current;
    if (!root) return;
    const sync = () => {
      setBarWidth(root.clientWidth);
      measure();
    };
    sync();
    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', sync);
      return () => window.removeEventListener('resize', sync);
    }
    const observer = new ResizeObserver(sync);
    observer.observe(root);
    return () => observer.disconnect();
  }, [measure, menuSignature]);

  // 文案变化（切换语言）后重新测量，避免沿用旧宽度
  useLayoutEffect(() => {
    widthCache.current = {};
    measure();
  }, [menuSignature, measure]);

  const closeAll = useCallback(() => {
    setOpenKey(null);
    setOpenSub(null);
  }, []);

  useEffect(() => {
    if (!openKey) return;
    const onPointerDown = (event: MouseEvent) => {
      if (!barRef.current?.contains(event.target as Node)) closeAll();
    };
    document.addEventListener('mousedown', onPointerDown);
    return () => document.removeEventListener('mousedown', onPointerDown);
  }, [openKey, closeAll]);

  useEffect(() => {
    if (!openKey) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      closeAll();
      barRef.current
        ?.querySelector<HTMLElement>(`[data-efk-menu-key="${openKey}"]`)
        ?.focus();
    };
    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
  }, [openKey, closeAll]);

  const widths = menus.map((menu) => widthCache.current[menu.key] ?? estimateWidth(menu.label));
  const allWidth = widths.reduce((sum, width, index) => sum + width + (index > 0 ? GAP : 0), 0);
  const moreWidth = 34;
  let visibleCount = menus.length;
  if (barWidth > 0 && allWidth > barWidth) {
    let used = 0;
    visibleCount = 0;
    for (let index = 0; index < menus.length; index += 1) {
      const next = used + (index > 0 ? GAP : 0) + widths[index];
      if (next + GAP + moreWidth > barWidth) break;
      used = next;
      visibleCount += 1;
    }
  }
  const visibleMenus = menus.slice(0, visibleCount);
  const hiddenMenus = menus.slice(visibleCount);
  const overflowing = hiddenMenus.length > 0;

  const visibleKeys = visibleMenus.map((menu) => menu.key).join('|');

  useEffect(() => {
    if (!openKey) return;
    if (openKey === MORE_KEY) {
      if (!overflowing) closeAll();
      return;
    }
    if (!visibleKeys.split('|').includes(openKey)) closeAll();
  }, [openKey, overflowing, visibleKeys, closeAll]);

  const handleSelect = (item: OfficeMenuItem) => {
    if (item.disabled) return;
    item.onSelect?.();
    closeAll();
  };

  const moreMenu: OfficeMenu = {
    key: MORE_KEY,
    label: t('shell.more'),
    align: 'end',
    items: [
      hiddenMenus.map((menu) => ({
        key: menu.key,
        label: menu.label,
        children: menu.items,
      })),
    ],
  };

  const renderDropdown = (menu: OfficeMenu, alignEnd: boolean) => (
    <div
      className={`efk-dropdown${alignEnd ? ' efk-dropdown--end' : ''}`}
      role="menu"
      aria-label={menu.label}
      onKeyDown={(event: ReactKeyboardEvent<HTMLDivElement>) => {
        if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
        const items = Array.from(
          event.currentTarget.querySelectorAll<HTMLButtonElement>(
            ':scope > .efk-menu-row > .efk-menu-item:not(:disabled)',
          ),
        );
        if (items.length === 0) return;
        event.preventDefault();
        const current = items.indexOf(document.activeElement as HTMLButtonElement);
        const step = event.key === 'ArrowDown' ? 1 : -1;
        const next = current < 0 ? (step === 1 ? 0 : items.length - 1) : (current + step + items.length) % items.length;
        items[next].focus();
      }}
    >
      <MenuRows
        groups={menu.items}
        openSub={openSub}
        setOpenSub={setOpenSub}
        onSelect={handleSelect}
        flipSubmenu={alignEnd}
      />
    </div>
  );

  return (
    <div
      ref={barRef}
      className={`efk-menu-bar${className ? ` ${className}` : ''}`}
      role="menubar"
      aria-label={ariaLabel ?? t('shell.menubar')}
    >
      {menus.map((menu, index) => {
        if (index >= visibleCount) return null;
        const alignEnd = (menu.align ?? (index >= menus.length - 2 ? 'end' : 'start')) === 'end';
        return (
          <div className="efk-menu-root" key={menu.key}>
            <button
              type="button"
              role="menuitem"
              aria-haspopup="menu"
              aria-expanded={openKey === menu.key}
              data-efk-menu-key={menu.key}
              className={`efk-menu-btn${openKey === menu.key ? ' is-open' : ''}`}
              onClick={() => {
                if (openKey === menu.key) closeAll();
                else {
                  setOpenKey(menu.key);
                  setOpenSub(null);
                }
              }}
              onMouseEnter={() => {
                if (openKey && openKey !== menu.key) {
                  setOpenKey(menu.key);
                  setOpenSub(null);
                }
              }}
            >
              {menu.label}
            </button>
            {openKey === menu.key && renderDropdown(menu, alignEnd)}
          </div>
        );
      })}
      {overflowing && (
        <div className="efk-menu-root">
          <button
            type="button"
            role="menuitem"
            aria-haspopup="menu"
            aria-expanded={openKey === MORE_KEY}
            aria-label={t('shell.more')}
            title={t('shell.more')}
            data-efk-menu-key={MORE_KEY}
            className={`efk-menu-btn efk-menu-btn--icon${openKey === MORE_KEY ? ' is-open' : ''}`}
            onClick={() => {
              if (openKey === MORE_KEY) closeAll();
              else {
                setOpenKey(MORE_KEY);
                setOpenSub(null);
              }
            }}
            onMouseEnter={() => {
              if (openKey && openKey !== MORE_KEY) {
                setOpenKey(MORE_KEY);
                setOpenSub(null);
              }
            }}
          >
            <MoreDots />
          </button>
          {openKey === MORE_KEY && renderDropdown(moreMenu, true)}
        </div>
      )}
    </div>
  );
}
