// @vitest-environment jsdom
import { describe, expect, it, vi, afterEach } from 'vitest';
import { pickFile } from './filePicker';

afterEach(() => {
  vi.restoreAllMocks();
  vi.useRealTimers();
});

describe('filePicker 文件选择', () => {
  it('change 事件带文件时返回该文件，input 随后移除', async () => {
    const promise = pickFile('.xlsx,.xlsm');
    const input = document.querySelector<HTMLInputElement>('input[type="file"]')!;
    expect(input.accept).toBe('.xlsx,.xlsm');
    const file = new File(['x'], '报表.xlsx');
    Object.defineProperty(input, 'files', { value: [file] });
    input.dispatchEvent(new Event('change'));
    await expect(promise).resolves.toBe(file);
    expect(document.querySelector('input[type="file"]')).toBeNull();
  });

  it('cancel 事件返回 null', async () => {
    const promise = pickFile('.xlsx');
    const input = document.querySelector<HTMLInputElement>('input[type="file"]')!;
    input.dispatchEvent(new Event('cancel'));
    await expect(promise).resolves.toBeNull();
  });

  it('settle 后移除 focus 兜底监听器，不会二次 settle', async () => {
    vi.useFakeTimers();
    const removeSpy = vi.spyOn(window, 'removeEventListener');
    const promise = pickFile('.xlsx');
    const input = document.querySelector<HTMLInputElement>('input[type="file"]')!;
    const file = new File(['x'], '报表.xlsx');
    Object.defineProperty(input, 'files', { value: [file] });
    input.dispatchEvent(new Event('change'));
    await expect(promise).resolves.toBe(file);
    // 正常 change 路径即解绑 focus 兜底监听器（防泄漏）
    expect(removeSpy).toHaveBeenCalledWith('focus', expect.any(Function));
    // 旧浏览器兜底路径（focus + 500ms 延迟）在已 settle 后不再产生副作用
    window.dispatchEvent(new Event('focus'));
    vi.advanceTimersByTime(600);
    expect(document.querySelector('input[type="file"]')).toBeNull();
  });
});
