/// <reference types="vite/client" />
import type { FlowsApi } from '../shared/ipc'

declare global {
  interface Window {
    flowsApi: FlowsApi
  }
}

export {}