/** 打开文件选择框（临时 input，不依赖 React ref）；用户取消返回 null */
export function pickFile(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    let settled = false;
    const finish = (file: File | null) => {
      if (settled) return;
      settled = true;
      window.removeEventListener('focus', onFocus); // 防 focus 兜底监听器泄漏
      input.remove();
      resolve(file);
    };
    // 兜底：旧浏览器无 cancel 事件，选择框关闭（窗口重新聚焦）且无文件时视为取消
    const onFocus = () => {
      window.setTimeout(() => finish(input.files?.[0] ?? null), 500);
    };
    input.addEventListener('change', () => finish(input.files?.[0] ?? null));
    input.addEventListener('cancel', () => finish(null));
    window.addEventListener('focus', onFocus);
    document.body.appendChild(input);
    input.click();
  });
}
