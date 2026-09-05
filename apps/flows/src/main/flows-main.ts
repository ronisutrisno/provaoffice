import { app, BrowserWindow, dialog, ipcMain, Menu, WebContentsView } from 'electron'
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { FlowDocument } from '../shared/ipc'
import { emptyFlow } from '../shared/ipc'
import { registerAiIpc } from './ai-ipc'
import {
  configureFlowsRuntime,
  flowsIsDirty,
  pendingByWc,
  runtime,
  setFlowsDirty,
  type RuntimePaths,
} from './session-state'

export { configureFlowsRuntime, flowsIsDirty }
export type { RuntimePaths }

const FLOW_EXT = 'flow'

function parseFlow(text: string): FlowDocument {
  try {
    const doc = JSON.parse(text) as FlowDocument
    if (doc && Array.isArray(doc.nodes) && Array.isArray(doc.edges)) return doc
  } catch {
    /* fall through */
  }
  return emptyFlow()
}

let flowsIpcRegistered = false

function registerFlowsIpc(): void {
  if (flowsIpcRegistered) return
  flowsIpcRegistered = true
  ipcMain.handle('flows:get-language', () => {
    try {
      const s = JSON.parse(
        readFileSync(join(app.getPath('userData'), 'app-settings.json'), 'utf-8'),
      ) as { language?: string; lang?: string }
      return s.language ?? s.lang ?? 'en'
    } catch {
      return 'en'
    }
  })
  ipcMain.handle('flows:get-theme', () => 'system')

  ipcMain.handle('flows:open', async (event) => {
    const win = BrowserWindow.fromWebContents(event.sender)
    const res = await dialog.showOpenDialog(win ?? undefined!, {
      filters: [{ name: 'PROVAOffice Flow', extensions: [FLOW_EXT] }],
      properties: ['openFile'],
    })
    if (res.canceled || !res.filePaths[0]) return null
    const path = res.filePaths[0]
    const doc = parseFlow(readFileSync(path, 'utf-8'))
    return { path, doc }
  })

  ipcMain.handle('flows:save', async (event, doc: FlowDocument) => {
    const win = BrowserWindow.fromWebContents(event.sender)
    const res = await dialog.showSaveDialog(win ?? undefined!, {
      defaultPath: `${doc.title || 'flow'}.${FLOW_EXT}`,
      filters: [{ name: 'PROVAOffice Flow', extensions: [FLOW_EXT] }],
    })
    if (res.canceled || !res.filePath) return null
    writeFileSync(res.filePath, JSON.stringify(doc, null, 2), 'utf-8')
    setFlowsDirty(event.sender.id, false)
    return { path: res.filePath }
  })

  ipcMain.handle('flows:save-as', async (event, doc: FlowDocument) => {
    const win = BrowserWindow.fromWebContents(event.sender)
    const res = await dialog.showSaveDialog(win ?? undefined!, {
      defaultPath: `${doc.title || 'flow'}.${FLOW_EXT}`,
      filters: [{ name: 'PROVAOffice Flow', extensions: [FLOW_EXT] }],
    })
    if (res.canceled || !res.filePath) return null
    writeFileSync(res.filePath, JSON.stringify(doc, null, 2), 'utf-8')
    setFlowsDirty(event.sender.id, false)
    return { path: res.filePath }
  })

  ipcMain.handle('flows:export-png', async (event, dataUrl: string) => {
    const win = BrowserWindow.fromWebContents(event.sender)
    const res = await dialog.showSaveDialog(win ?? undefined!, {
      defaultPath: 'flow.png',
      filters: [{ name: 'PNG', extensions: ['png'] }],
    })
    if (res.canceled || !res.filePath) return null
    const b64 = dataUrl.replace(/^data:image\/png;base64,/, '')
    writeFileSync(res.filePath, Buffer.from(b64, 'base64'))
    return { path: res.filePath }
  })

  ipcMain.handle('flows:export-svg', async (event, svg: string) => {
    const win = BrowserWindow.fromWebContents(event.sender)
    const res = await dialog.showSaveDialog(win ?? undefined!, {
      defaultPath: 'flow.svg',
      filters: [{ name: 'SVG', extensions: ['svg'] }],
    })
    if (res.canceled || !res.filePath) return null
    writeFileSync(res.filePath, svg, 'utf-8')
    return { path: res.filePath }
  })

  ipcMain.handle('flows:consume-pending-open', (event) => {
    const id = event.sender.id
    const path = pendingByWc.get(id)
    if (!path) return null
    pendingByWc.delete(id)
    if (!existsSync(path)) return null
    return { path, doc: parseFlow(readFileSync(path, 'utf-8')) }
  })

  ipcMain.on('flows:dirty', (event, dirty: boolean) => {
    setFlowsDirty(event.sender.id, dirty)
  })
}

/** Tab version: hosted in the shell's WebContentsView. */
export function createFlowsView(openPath?: string | null): WebContentsView {
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  registerFlowsIpc()
  if (openPath && existsSync(openPath)) pendingByWc.set(view.webContents.id, openPath)
  if (runtime.rendererDevUrl) {
    const devUrl = new URL(runtime.rendererDevUrl)
    devUrl.searchParams.set('mode', 'tab')
    void view.webContents.loadURL(devUrl.toString())
  } else if (runtime.rendererFilePath) {
    void view.webContents.loadFile(runtime.rendererFilePath, { query: { mode: 'tab' } })
  }
  return view
}

export function createFlowsWindow(openPath?: string | null): BrowserWindow {
  const win = new BrowserWindow({
    width: 1280,
    height: 840,
    title: 'PROVAOffice Flows',
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  registerFlowsIpc()
  registerAiIpc()
  if (openPath && existsSync(openPath)) pendingByWc.set(win.webContents.id, openPath)
  if (runtime.rendererDevUrl) win.loadURL(runtime.rendererDevUrl)
  else if (runtime.rendererFilePath) win.loadFile(runtime.rendererFilePath)
  return win
}

export function buildFlowsMenu(): Menu {
  const template: Electron.MenuItemConstructorOptions[] = [
    {
      label: 'File',
      submenu: [
        { label: 'New Flow', click: () => send('new') },
        { label: 'Open…', click: () => send('open') },
        { type: 'separator' },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => send('save') },
        { label: 'Save As…', click: () => send('save-as') },
        { type: 'separator' },
        { label: 'Export PNG…', click: () => send('export-png') },
        { label: 'Export SVG…', click: () => send('export-svg') },
      ],
    },
    { role: 'editMenu' },
    { role: 'viewMenu' },
  ]
  return Menu.buildFromTemplate(template)
}

function send(cmd: string): void {
  const win = BrowserWindow.getFocusedWindow()
  win?.webContents.send('flows:menu', cmd)
}

export function startFlowsStandalone(): void {
  app.whenReady().then(() => {
    configureFlowsRuntime({
      preloadPath: runtime.preloadPath,
      rendererDevUrl: process.env.ELECTRON_RENDERER_URL,
      rendererFilePath: runtime.rendererFilePath,
    })
    Menu.setApplicationMenu(buildFlowsMenu())
    createFlowsWindow()
    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createFlowsWindow()
    })
  })
}
