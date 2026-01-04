import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
      '@insight/license': path.resolve(__dirname, './insight-common/license/typescript/index.ts'),
    },
  },
  // Tauri expects a fixed port
  server: {
    port: 1420,
    strictPort: true,
  },
  // Tauri CLI will handle the build output
  build: {
    target: ['es2021', 'chrome100', 'safari13'],
    minify: !process.env.TAURI_DEBUG ? 'esbuild' : false,
    sourcemap: !!process.env.TAURI_DEBUG,
  },
});
