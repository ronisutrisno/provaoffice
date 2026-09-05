import { join } from 'node:path'

export interface RuntimePaths {
  preloadPath: string
  rendererDevUrl?: string | undefined
  rendererFilePath?: string | undefined
}

export const runtime: RuntimePaths = {
  preloadPath: join(__dirname, '../preload/index.js'),
  rendererDevUrl: process.env.ELECTRON_RENDERER_URL,
  rendererFilePath: join(__dirname, '../renderer/index.html'),
}

export function configureFlowsRuntime(paths: RuntimePaths): void {
  runtime.preloadPath = paths.preloadPath
  runtime.rendererDevUrl = paths.rendererDevUrl
  runtime.rendererFilePath = paths.rendererFilePath
}

/** Dirty state per webContents (tab mode close guard). */
const dirtyByWc = new Map<number, boolean>()

export function setFlowsDirty(id: number, dirty: boolean): void {
  dirtyByWc.set(id, dirty)
}

export function flowsIsDirty(id: number): boolean {
  return dirtyByWc.get(id) ?? false
}

/** Pending open path per webContents (tab mode: renderer consumes after mount). */
export const pendingByWc = new Map<number, string>()
