import { contextBridge, ipcRenderer } from 'electron'
import type { AiSettings, FlowsApi } from '../shared/ipc'

const api: FlowsApi = {
  getLanguage: () => ipcRenderer.invoke('flows:get-language'),
  getTheme: () => ipcRenderer.invoke('flows:get-theme'),
  onThemeChanged(cb) {
    const listener = (_e: unknown, theme: 'system' | 'light' | 'dark'): void => cb(theme)
    ipcRenderer.on('flows:theme-changed', listener)
    return () => ipcRenderer.removeListener('flows:theme-changed', listener)
  },
  newFlow: () => ipcRenderer.invoke('home:new-flow'),
  openFlow: () => ipcRenderer.invoke('flows:open'),
  saveFlow: (doc) => ipcRenderer.invoke('flows:save', doc),
  saveFlowAs: (doc) => ipcRenderer.invoke('flows:save-as', doc),
  exportPng: (dataUrl) => ipcRenderer.invoke('flows:export-png', dataUrl),
  exportSvg: (svg) => ipcRenderer.invoke('flows:export-svg', svg),
  consumePendingOpen: () => ipcRenderer.invoke('flows:consume-pending-open'),
  onMenu(cb) {
    const listener = (_e: unknown, cmd: string): void => cb(cmd)
    ipcRenderer.on('flows:menu', listener)
    return () => ipcRenderer.removeListener('flows:menu', listener)
  },
  aiStream: (request) => ipcRenderer.invoke('ai:stream', request),
  onAiStream(cb) {
    const listener = (_e: unknown, chunk: Parameters<typeof cb>[0]): void => cb(chunk)
    ipcRenderer.on('ai:stream-chunk', listener)
    return () => ipcRenderer.removeListener('ai:stream-chunk', listener)
  },
  aiStreamCancel: (requestId) => ipcRenderer.invoke('ai:stream-cancel', requestId),
  getAiSettings: () => ipcRenderer.invoke('ai:get-settings'),
  setAiSettings: (settings: AiSettings) => ipcRenderer.invoke('ai:set-settings', settings),
  pickAttachments: () => ipcRenderer.invoke('files:pick'),
  readAttachment: (path) => ipcRenderer.invoke('files:read', path),
  readAttachmentImage: (path) => ipcRenderer.invoke('files:read-image', path),
  webSearch: (query, maxResults) => ipcRenderer.invoke('ai:web-search', query, maxResults),
  imageSearch: (query, maxResults) => ipcRenderer.invoke('ai:image-search', query, maxResults),
}

contextBridge.exposeInMainWorld('flowsApi', api)