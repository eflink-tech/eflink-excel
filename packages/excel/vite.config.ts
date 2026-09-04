import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import { defineConfig } from 'vite';

// 库构建：ESM + 类型（tsc 单独产出）+ 单文件 styles.css（含 tailwind 工具类与 Univer 样式）
export default defineConfig({
  plugins: [react(), tailwindcss()],
  build: {
    lib: {
      entry: 'src/index.ts',
      formats: ['es'],
      fileName: () => 'index.js',
    },
    // watch 模式（pnpm dev）不清空 dist，保留 tsc 产出的 dist/types；生产构建经 --emptyOutDir 显式清空
    emptyOutDir: false,
    cssCodeSplit: false,
    rollupOptions: {
      external: (id) =>
        /^(react|react-dom|react\/jsx-runtime|@univerjs\/|dexie|echarts|exceljs|html2canvas|lucide-react|zustand)/.test(id),
      output: {
        assetFileNames: (asset) =>
          asset.names?.[0]?.endsWith('.css') ? 'styles.css' : 'assets/[name][extname]',
      },
    },
  },
});
