import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

const here = dirname(fileURLToPath(import.meta.url))

const workspaceAlias = {
  '@prova/ai-search': resolve(here, '../../packages/ai-search/src/index.ts'),
  '@prova/ai-provider': resolve(here, '../../packages/ai-provider/src/index.ts'),
  '@prova/agent-core': resolve(here, '../../packages/agent-core/src/index.ts'),
  '@prova/electron-utils': resolve(here, '../../packages/electron-utils/src/index.ts'),
  '@prova/file-parse': resolve(here, '../../packages/file-parse/src/index.ts'),
  '@prova/i18n': resolve(here, '../../packages/i18n/src/index.ts'),
}

export default defineConfig({
  main: {
    resolve: { alias: workspaceAlias },
    plugins: [
      externalizeDepsPlugin({
        exclude: ['@prova/ai-search', '@prova/ai-provider', '@prova/electron-utils', '@prova/file-parse'],
      }),
    ],
  },
  preload: {
    plugins: [externalizeDepsPlugin()],
  },
  renderer: {
    resolve: { alias: workspaceAlias },
    plugins: [react()],
    server: {
      port: Number(process.env.FLOWS_DEV_PORT) || 5178,
      strictPort: Boolean(process.env.FLOWS_DEV_PORT),
    },
  },
})
