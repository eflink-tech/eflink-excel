import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    environment: 'node',
    setupFiles: ['./src/test/setup.ts'],
    exclude: ['**/node_modules/**', 'e2e/**'],
  },
});
