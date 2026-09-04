import { create } from 'zustand';

/** 确认弹窗配置 */
export interface ConfirmOptions {
  title: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
}

interface PendingConfirm extends ConfirmOptions {
  resolve: (ok: boolean) => void;
}

interface UiStore {
  toast: string | null;
  showToast: (msg: string) => void;
  clearToast: () => void;
  /** 当前打开「图表设置」面板的图片 id（null = 面板关闭） */
  chartPanelDrawingId: string | null;
  openChartPanel: (drawingId: string) => void;
  closeChartPanel: () => void;
  /** 待确认的弹窗（null = 关闭）；requestConfirm 发起，ConfirmDialogHost 消费 */
  confirm: PendingConfirm | null;
  /** 发起确认：resolve(true)=确定，resolve(false)=取消 */
  requestConfirm: (opts: ConfirmOptions) => Promise<boolean>;
  resolveConfirm: (ok: boolean) => void;
}

export const useUiStore = create<UiStore>((set) => ({
  toast: null,
  showToast: (msg) => set({ toast: msg }),
  clearToast: () => set({ toast: null }),
  chartPanelDrawingId: null,
  openChartPanel: (drawingId) => set({ chartPanelDrawingId: drawingId }),
  closeChartPanel: () => set({ chartPanelDrawingId: null }),
  confirm: null,
  requestConfirm: (opts) =>
    new Promise<boolean>((resolve) => set({ confirm: { ...opts, resolve } })),
  resolveConfirm: (ok) => {
    set((s) => {
      s.confirm?.resolve(ok);
      return { confirm: null };
    });
  },
}));
