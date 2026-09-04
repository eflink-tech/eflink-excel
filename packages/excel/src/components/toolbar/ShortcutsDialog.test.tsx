// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen } from '@testing-library/react';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { ShortcutsDialog } from './ShortcutsDialog';

afterEach(() => {
  cleanup();
});

function openDialog(onClose = vi.fn()) {
  render(<ShortcutsDialog open onClose={onClose} />);
  return onClose;
}

describe('ShortcutsDialog 快捷键弹窗', () => {
  it('open=false 时不渲染', () => {
    render(<ShortcutsDialog open={false} onClose={() => {}} />);
    expect(screen.queryByText('快捷键')).toBeNull();
  });

  it('打开后展示分组与快捷键条目', () => {
    openDialog();
    expect(screen.getByText('快捷键')).toBeTruthy();
    expect(screen.getByText('保存到本地数据库')).toBeTruthy();
    expect(screen.getByText('⌘S')).toBeTruthy();
    expect(screen.getByText('单元格内换行')).toBeTruthy();
    expect(screen.getByText('扩展选区')).toBeTruthy();
  });

  it('按 Escape 关闭并回调 onClose', () => {
    const onClose = openDialog();
    fireEvent.keyDown(window, { key: 'Escape' });
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it('点击关闭按钮回调 onClose', () => {
    const onClose = openDialog();
    fireEvent.click(screen.getByRole('button', { name: '关闭' }));
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it('点击遮罩关闭，点击内容区不关闭', () => {
    const onClose = vi.fn();
    render(<ShortcutsDialog open onClose={onClose} />);
    const overlay = document.body.querySelector('.shortcuts-dialog-overlay') as HTMLElement;
    const panel = overlay.querySelector('.shortcuts-dialog') as HTMLElement;
    fireEvent.click(panel);
    expect(onClose).not.toHaveBeenCalled();
    fireEvent.click(overlay);
    expect(onClose).toHaveBeenCalledTimes(1);
  });
});
