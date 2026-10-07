import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/bugu-kanri/',
  define: {
    'import.meta.env.VITE_DISCORD_WEBHOOK_URL': JSON.stringify(process.env.VITE_DISCORD_WEBHOOK_URL)
  }
})
