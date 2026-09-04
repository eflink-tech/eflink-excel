import { useRef } from 'react';
import { MenuDropdown } from './MenuDropdown';
import './myfsheet.css';

/** 顶栏 ☰ 图标菜单按钮 */
export function AppMenu() {
  const importInputRef = useRef<HTMLInputElement>(null);

  return (
    <>
      <input
        ref={importInputRef}
        type="file"
        accept=".efexcel,application/json"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) import('./../../core/fileActions').then(({ importFileAction }) => { void importFileAction(file) });
          e.target.value = '';
        }}
      />
      <MenuDropdown
        trigger={
          <button
            aria-label="主菜单"
            title="主菜单"
            className="flex h-7 w-7 items-center justify-center rounded-md text-slate-600 hover:bg-slate-100"
          >
            <span className="myf-icon myf-icon-file-menu" />
          </button>
        }
        importInputRef={importInputRef}
      />
    </>
  );
}
