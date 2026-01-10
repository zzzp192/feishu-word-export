import { defineConfig } from 'vite';

export default defineConfig({
  base: './',
  server: {
    host: '0.0.0.0',
  },
  // --- 新增下面这段 ---
  build: {
    outDir: 'docs', // 显式指定输出目录为 docs
  }
})
