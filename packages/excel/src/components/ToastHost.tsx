import { useEffect } from 'react';
import { useUiStore } from '../store/uiStore';

const TOAST_MS = 2500;

export function ToastHost() {
  const toast = useUiStore((s) => s.toast);
  const clearToast = useUiStore((s) => s.clearToast);

  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(clearToast, TOAST_MS);
    return () => clearTimeout(timer);
  }, [toast, clearToast]);

  if (!toast) return null;
  return (
    <div className="pointer-events-none fixed left-1/2 top-6 z-[9999] -translate-x-1/2">
      <div className="rounded bg-slate-800/90 px-4 py-2 text-sm text-white shadow-lg">{toast}</div>
    </div>
  );
}
