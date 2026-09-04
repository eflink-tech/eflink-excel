// @vitest-environment jsdom
import { act, cleanup, render, screen } from '@testing-library/react';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { ToastHost } from './ToastHost';
import { useUiStore } from '../store/uiStore';

beforeEach(() => {
  vi.useFakeTimers();
  useUiStore.getState().clearToast();
});

afterEach(() => {
  vi.useRealTimers();
  cleanup();
});

describe('ToastHost 全局提示条', () => {
  it('无 toast 时不渲染任何内容', () => {
    const { container } = render(<ToastHost />);
    expect(container.textContent).toBe('');
  });

  it('showToast 后渲染提示文案', () => {
    render(<ToastHost />);
    act(() => {
      useUiStore.getState().showToast('已保存');
    });
    expect(screen.getByText('已保存')).toBeTruthy();
  });

  it('2500ms 后自动消失', () => {
    render(<ToastHost />);
    act(() => {
      useUiStore.getState().showToast('已导出 xlsx');
    });
    act(() => {
      vi.advanceTimersByTime(2499);
    });
    expect(screen.getByText('已导出 xlsx')).toBeTruthy();
    act(() => {
      vi.advanceTimersByTime(1);
    });
    expect(screen.queryByText('已导出 xlsx')).toBeNull();
  });

  it('新 toast 替换旧 toast 并重置消失计时', () => {
    render(<ToastHost />);
    act(() => {
      useUiStore.getState().showToast('第一条');
    });
    act(() => {
      vi.advanceTimersByTime(2000);
    });
    act(() => {
      useUiStore.getState().showToast('第二条');
    });
    act(() => {
      // 若计时未重置，第一条的 2500ms 已到期；这里验证只按第二条计时
      vi.advanceTimersByTime(2499);
    });
    expect(screen.getByText('第二条')).toBeTruthy();
    act(() => {
      vi.advanceTimersByTime(1);
    });
    expect(screen.queryByText('第二条')).toBeNull();
  });
});
