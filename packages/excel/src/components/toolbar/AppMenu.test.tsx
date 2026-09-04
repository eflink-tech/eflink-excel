// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen } from '@testing-library/react';
import { afterEach, describe, expect, it, vi } from 'vitest';

vi.mock('../../core/univer/commands', () => ({
  undo: vi.fn(),
  redo: vi.fn(),
  clearContentOnly: vi.fn(),
  clearFormatOnly: vi.fn(),
  freezeFirstRow: vi.fn(),
  freezeFirstColumn: vi.fn(),
  cancelFreeze: vi.fn(),
}));

vi.mock('../../core/fileActions', () => ({
  exportPngAction: vi.fn(),
  exportEfexcelAction: vi.fn(),
  importFileAction: vi.fn(),
  newDocAction: vi.fn(),
  saveAction: vi.fn(),
}));

import * as cmd from '../../core/univer/commands';
import * as actions from '../../core/fileActions';
import { AppMenu } from './AppMenu';

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

function openMenu() {
  cleanup();
  render(<AppMenu />);
  fireEvent.click(screen.getByRole('button', { name: '主菜单' }));
}

function branch(label: string) {
  return screen.getByRole('button', { name: label });
}

describe('AppMenu 主菜单', () => {
  it('点击 ☰ 展开一级面板（文件/编辑/查看/快捷键）', () => {
    openMenu();
    expect(branch('文件')).toBeTruthy();
    expect(branch('编辑')).toBeTruthy();
    expect(branch('查看')).toBeTruthy();
    expect(branch('快捷键')).toBeTruthy();
  });

  it('悬停「文件」展开二级面板并触发导出动作', () => {
    openMenu();
    fireEvent.mouseEnter(branch('文件'));
    expect(screen.getByRole('button', { name: '新建表格' })).toBeTruthy();
    expect(screen.getByRole('button', { name: '保存⌘S' })).toBeTruthy();
    expect(screen.getByRole('button', { name: '导入表格(.efexcel)' })).toBeTruthy();
    expect(screen.getByRole('button', { name: '导出表格(.efexcel)' })).toBeTruthy();
    expect(screen.getByRole('button', { name: '导出图片' })).toBeTruthy();

    fireEvent.click(screen.getByRole('button', { name: '导出表格(.efexcel)' }));
    expect(actions.exportEfexcelAction).toHaveBeenCalledTimes(1);
    // 执行动作后面板收起
    expect(screen.queryByRole('button', { name: '新建表格' })).toBeNull();
  });

  it('编辑面板：撤销 / 重做 / 清除内容 / 清除格式', () => {
    openMenu();
    fireEvent.mouseEnter(branch('编辑'));
    fireEvent.click(screen.getByRole('button', { name: '撤销' }));
    expect(cmd.undo).toHaveBeenCalledTimes(1);

    openMenu();
    fireEvent.mouseEnter(branch('编辑'));
    fireEvent.click(screen.getByRole('button', { name: '重做' }));
    expect(cmd.redo).toHaveBeenCalledTimes(1);

    openMenu();
    fireEvent.mouseEnter(branch('编辑'));
    fireEvent.click(screen.getByRole('button', { name: '清除内容' }));
    expect(cmd.clearContentOnly).toHaveBeenCalledTimes(1);

    openMenu();
    fireEvent.mouseEnter(branch('编辑'));
    fireEvent.click(screen.getByRole('button', { name: '清除格式' }));
    expect(cmd.clearFormatOnly).toHaveBeenCalledTimes(1);
  });

  it('查看面板：冻结首行 / 冻结首列 / 取消冻结', () => {
    openMenu();
    fireEvent.mouseEnter(branch('查看'));
    fireEvent.click(screen.getByRole('button', { name: '冻结首行' }));
    expect(cmd.freezeFirstRow).toHaveBeenCalledTimes(1);

    openMenu();
    fireEvent.mouseEnter(branch('查看'));
    fireEvent.click(screen.getByRole('button', { name: '冻结首列' }));
    expect(cmd.freezeFirstColumn).toHaveBeenCalledTimes(1);

    openMenu();
    fireEvent.mouseEnter(branch('查看'));
    fireEvent.click(screen.getByRole('button', { name: '取消冻结' }));
    expect(cmd.cancelFreeze).toHaveBeenCalledTimes(1);
  });

  it('「快捷键」叶子项直接打开快捷键弹窗', () => {
    openMenu();
    fireEvent.click(branch('快捷键'));
    expect(screen.getByText('快捷键')).toBeTruthy();
    expect(screen.getByText('保存到本地数据库')).toBeTruthy();
  });

  it('点击菜单外空白处关闭（capture 阶段监听 Univer 画布场景）', () => {
    openMenu();
    expect(branch('文件')).toBeTruthy();
    fireEvent.click(document.body);
    expect(screen.queryByRole('button', { name: '文件' })).toBeNull();
  });

  it('保存动作触发 saveAction', () => {
    openMenu();
    fireEvent.mouseEnter(branch('文件'));
    fireEvent.click(screen.getByRole('button', { name: '保存⌘S' }));
    expect(actions.saveAction).toHaveBeenCalledTimes(1);
  });
});
