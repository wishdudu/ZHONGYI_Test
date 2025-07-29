import { fileURLToPath, URL } from 'node:url'

import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import vueJsx from '@vitejs/plugin-vue-jsx'
import vueDevTools from 'vite-plugin-vue-devtools'
import path from 'node:path'

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    vue(),
    vueJsx(),
    vueDevTools(),
  ],
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url))
    },
  },
  server: {
    proxy: {
      '/zwx': {
        target: 'http://60.12.208.134',
        changeOrigin: true,
        rewrite: path => path.replace(/^\/zwx/, '')
      },
      '/baidu': {
        target: 'https://aip.baidubce.com',
        changeOrigin: true,
        secure: false, 
        rewrite: path => path.replace(/^\/baidu/, '')
      }
    }
  }
})
