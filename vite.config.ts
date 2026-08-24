import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import {defineConfig, loadEnv} from 'vite';

export default defineConfig(({mode}) => {
  const env = loadEnv(mode, '.', '');
  return {
    base: '/',
    plugins: [react(), tailwindcss()],
    define: {
      'process.env.GEMINI_API_KEY': JSON.stringify(env.GEMINI_API_KEY || ''),
      'import.meta.env.MONGODB_URI': JSON.stringify(env.MONGODB_URI || ''),
      'import.meta.env.MONGODB_DB_NAME': JSON.stringify(env.MONGODB_DB_NAME || ''),
      'import.meta.env.VITE_GITHUB_EXCEL_URL': JSON.stringify(env.VITE_GITHUB_EXCEL_URL || ''),
      'import.meta.env.HOSTINGER_DB_HOST': JSON.stringify(env.HOSTINGER_DB_HOST || ''),
      'import.meta.env.HOSTINGER_DB_USER': JSON.stringify(env.HOSTINGER_DB_USER || ''),
      'import.meta.env.HOSTINGER_DB_NAME': JSON.stringify(env.HOSTINGER_DB_NAME || ''),
      'import.meta.env.HOSTINGER_DB_PORT': JSON.stringify(env.HOSTINGER_DB_PORT || ''),
    },
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modifyâfile watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
    },
  };
});
