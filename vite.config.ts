import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { getHttpsServerOptions } from 'office-addin-dev-certs'

// https://vite.dev/config/
export default defineConfig(async () => ({
  plugins: [react()],
  server: {
    port: 3000,
    https: await getHttpsServerOptions(),
  },
  build: {
    rollupOptions: {
      input: {
        main: 'index.html',
        taskpane: 'src/word-addin/taskpane.html',
      },
    },
  },
}))
