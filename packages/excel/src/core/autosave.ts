// 防抖调度器：延迟 N 秒执行回调（当前用于内容变化后写本地草稿），flush 立即触发一次；与 store 解耦便于测试

export interface Autosaver {
  /** 内容变化时调用：重置 N 秒倒计时 */
  schedule(): void;
  /** 立即执行回调（若定时器在等待则合并为一次） */
  flush(): Promise<void>;
  cancel(): void;
}

export function createAutosaver(save: () => void | Promise<void>, delayMs = 1500): Autosaver {
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
