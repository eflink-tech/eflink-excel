// @vitest-environment jsdom
import { describe, expect, it, vi, afterEach } from 'vitest';
import { pickFile } from './filePicker';

afterEach(() => {
  vi.restoreAllMocks();
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
});
