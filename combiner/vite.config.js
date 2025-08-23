import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    react(),
    tailwindcss(),
  ],
  server: {
    mimeTypes: {
      'application/wasm': ['wasm']
    }
  },
  assetsInclude: ["**/*.zip"],
  base: "regex-combiner",
  build: {
    rollupOptions: {
      input: {
        main: "index.html"
      }
    }
  }
})
