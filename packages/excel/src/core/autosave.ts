// 防抖自动保存 + 立即保存（Ctrl+S / 切后台时 flush），与 store 解耦便于测试

export interface Autosaver {
  /** 内容变化时调用：重置 N 秒倒计时 */
  schedule(): void;
  /** 立即保存（若定时器在等待则合并为一次） */
  flush(): Promise<void>;
  cancel(): void;
}

export function createAutosaver(save: () => Promise<void>, delayMs = 1500): Autosaver {
  let timer: ReturnType<typeof setTimeout> | null = null;
  let running: Promise<void> = Promise.resolve();

  const run = () => {
    running = running.then(save, () => {});
  };

  return {
    schedule() {
      if (timer) clearTimeout(timer);
      timer = setTimeout(() => {
        timer = null;
        run();
      }, delayMs);
    },
    async flush() {
      if (timer) {
        clearTimeout(timer);
        timer = null;
      }
      run();
      await running;
    },
    cancel() {
      if (timer) {
        clearTimeout(timer);
        timer = null;
      }
    },
  };
}
