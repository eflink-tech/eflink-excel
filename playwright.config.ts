import { defineConfig } from '@playwright/test';

// 注意用 localhost 而非 127.0.0.1：vite 8 默认可能只绑 IPv6 ::1
// 端口用 5179：5173 常被本机其他 eflink 子项目（word 等）的 dev server 占用
export default defineConfig({
  testDir: './e2e',
  use: {
    baseURL: 'http://localhost:5179',
  },
  webServer: {
    // e2e 跑生产构建（vite preview）：dev 实时编译整库在 CI 弱机上过慢，preview 更快更稳定
    command: 'pnpm --filter eflink-excel-demo build && pnpm --filter eflink-excel-demo exec vite preview --port 5179 --strictPort',
    url: 'http://localhost:5179',
    reuseExistingServer: true,
  },
});
