import { useEffect } from 'react';
import { useUiStore } from '../store/uiStore';

/**
 * 确认弹窗宿主：SheetEditor 内部挂载，经 useUiStore.requestConfirm(opts) 发起，
 * resolve(true)=确定 / resolve(false)=取消（含遮罩点击与 Escape）。
 */
export function ConfirmDialogHost() {
  const confirm = useUiStore((s) => s.confirm);
  const resolveConfirm = useUiStore((s) => s.resolveConfirm);

  useEffect(() => {
    if (!confirm) return;
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') resolveConfirm(false);
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [confirm, resolveConfirm]);

  if (!confirm) return null;
  return (
    <div
      className="fixed inset-0 z-[10000] flex items-center justify-center bg-black/45"
      onMouseDown={() => resolveConfirm(false)}
    >
      <div
        role="dialog"
        aria-label={confirm.title}
        className="w-[400px] max-w-[calc(100vw-48px)] rounded-xl bg-white p-6 shadow-2xl"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <h3 className="text-lg font-semibold text-slate-900">{confirm.title}</h3>
        <p className="mt-3 text-sm leading-6 text-slate-500">{confirm.message}</p>
        <div className="mt-6 flex justify-end gap-3">
          <button
            className="rounded-md border border-slate-200 bg-white px-4 py-1.5 text-sm text-slate-700 hover:bg-slate-50"
            onClick={() => resolveConfirm(false)}
          >
            {confirm.cancelText ?? '取消'}
          </button>
          <button
            className="rounded-md bg-[#2b7fff] px-4 py-1.5 text-sm text-white hover:bg-[#1f6ff2]"
            onClick={() => resolveConfirm(true)}
          >
            {confirm.confirmText ?? '确定'}
          </button>
        </div>
      </div>
    </div>
  );
}
