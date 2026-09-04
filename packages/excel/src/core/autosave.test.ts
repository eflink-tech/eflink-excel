import { afterEach, describe, expect, it, vi } from 'vitest';
import { createAutosaver } from './autosave';

afterEach(() => {
  vi.useRealTimers();
});

describe('防抖自动保存', () => {
  it('schedule 后到延迟时间执行一次保存', async () => {
    vi.useFakeTimers();
    let saves = 0;
    const saver = createAutosaver(async () => {
      saves += 1;
    }, 1000);
    saver.schedule();
    await vi.advanceTimersByTimeAsync(999);
    expect(saves).toBe(0);
    await vi.advanceTimersByTimeAsync(1);
    expect(saves).toBe(1);
  });

  it('连续 schedule 合并为一次保存', async () => {
    vi.useFakeTimers();
    let saves = 0;
    const saver = createAutosaver(async () => {
      saves += 1;
    }, 1000);
    saver.schedule();
    await vi.advanceTimersByTimeAsync(600);
    saver.schedule();
    await vi.advanceTimersByTimeAsync(600);
    expect(saves).toBe(0);
    await vi.advanceTimersByTimeAsync(400);
    expect(saves).toBe(1);
  });

  it('flush 立即保存，且不与待执行的定时器重复', async () => {
    vi.useFakeTimers();
    let saves = 0;
    const saver = createAutosaver(async () => {
      saves += 1;
    }, 1000);
    saver.schedule();
    await saver.flush();
    expect(saves).toBe(1);
    await vi.advanceTimersByTimeAsync(5000);
    expect(saves).toBe(1);
  });
});
