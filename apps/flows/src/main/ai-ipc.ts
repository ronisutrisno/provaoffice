import { app, BrowserWindow, dialog, ipcMain, net, shell } from 'electron'
import { readFileSync, statSync, writeFileSync } from 'node:fs'
import { basename, extname } from 'node:path'
import { join } from 'node:path'
import { parseFileToText } from '@prova/file-parse'
import {
  defaultAiSettings,
  isAiNetworkError,
  resolveAiSettings,
  setRescueFetch,
  streamForProvider,
  AiCreditsError,
  AiTimeoutError,
  type AiSettings,
  type AiStreamChunk,
  type AiStreamRequest,
  type LegacyAiSettings,
} from '@prova/ai-provider'
import { imageSearch, webSearch } from '@prova/ai-search'
import { ATTACHMENT_IMAGE_EXTS, type AttachmentAddResult, type AttachmentMeta } from '../shared/ipc'

const AI_SETTINGS_PATH = (): string => join(app.getPath('userData'), 'ai-settings.json')

function readJson<T>(path: string, fallback: T): T {
  try {
    return JSON.parse(readFileSync(path, 'utf-8')) as T
  } catch {
    return fallback
  }
}

function writeJson(path: string, value: unknown): void {
  writeFileSync(path, JSON.stringify(value, null, 2), 'utf-8')
}

const activeAiStreams = new Map<string, AbortController>()

let registered = false

export function registerAiIpc(): void {
  if (registered) return
  registered = true
  setRescueFetch((url, init) => net.fetch(url, init))

  ipcMain.handle('ai:get-settings', (): AiSettings => {
    const stored = readJson<Partial<AiSettings> & LegacyAiSettings>(AI_SETTINGS_PATH(), {})
    const settings = resolveAiSettings(stored, defaultAiSettings())
    settings.provider = 'prova'
    return settings
  })

  ipcMain.handle('ai:set-settings', (_event, settings: AiSettings) => {
    writeJson(AI_SETTINGS_PATH(), settings)
  })

  ipcMain.handle('ai:stream', async (event, request: AiStreamRequest) => {
    const { requestId, settings, system, messages } = request
    const tools = request.tools ?? []
    const maxTokens = request.maxTokens ?? 8192
    const provider = settings.provider
    const config = settings.providers?.[provider]
    const send = (chunk: AiStreamChunk): void => {
      if (!event.sender.isDestroyed()) event.sender.send('ai:stream-chunk', chunk)
    }
    if (!config?.apiKey) {
      send({ requestId, type: 'error', error: 'API key required — set it in AI Provider Settings' })
      return
    }
    if (!config.model) {
      send({ requestId, type: 'error', error: 'No model selected' })
      return
    }
    const controller = new AbortController()
    activeAiStreams.set(requestId, controller)
    let lastPing = 0
    const ping = (): void => {
      const now = Date.now()
      if (now - lastPing < 5_000) return
      lastPing = now
      send({ requestId, type: 'ping' })
    }
    try {
      await streamForProvider(provider, config, system, messages, tools, maxTokens, {
        signal: controller.signal,
        onDelta: (text) => send({ requestId, type: 'delta', text }),
        onToolCall: (toolCall) => send({ requestId, type: 'tool-call', toolCall }),
        onActivity: ping,
      })
      send({ requestId, type: 'done' })
    } catch (err) {
      if (controller.signal.aborted) {
        send({ requestId, type: 'done' })
      } else {
        const msg = err instanceof Error ? err.message : String(err)
        console.error(`[flows ai-stream] ${requestId} failed:`, msg)
        send({
          requestId,
          type: 'error',
          error: msg,
          ...(err instanceof AiTimeoutError
            ? { errorCode: 'timeout' as const }
            : err instanceof AiCreditsError
              ? { errorCode: 'credits' as const }
              : isAiNetworkError(err)
                ? { errorCode: 'network' as const }
                : {}),
        })
      }
    } finally {
      activeAiStreams.delete(requestId)
    }
  })

  ipcMain.handle('ai:stream-cancel', (_event, requestId: string) => {
    activeAiStreams.get(requestId)?.abort()
  })

  ipcMain.handle('ai:web-search', async (_event, query: string, maxResults?: number) => {
    try {
      return await webSearch(String(query), typeof maxResults === 'number' ? maxResults : 6)
    } catch (err) {
      return { results: [], method: 'error', error: String(err) }
    }
  })

  ipcMain.handle('ai:image-search', async (_event, query: string, maxResults?: number) => {
    try {
      return await imageSearch(String(query), typeof maxResults === 'number' ? maxResults : 8)
    } catch (err) {
      return { images: [], method: 'error', error: String(err) }
    }
  })

  const ATTACH_EXTS = new Set(['txt', 'md', 'markdown', 'csv', 'tsv', 'json', 'yaml', 'yml', 'xml', 'html', 'htm', 'log', 'docx', 'pdf', 'pptx', 'xlsx', 'xls', 'png', 'jpg', 'jpeg', 'gif', 'webp'])
  const ATTACH_MAX = 50 * 1024 * 1024

  ipcMain.handle('files:pick', async (event): Promise<AttachmentAddResult | null> => {
    const win = BrowserWindow.fromWebContents(event.sender)
    const res = await dialog.showOpenDialog(win!, {
      properties: ['openFile', 'multiSelections'],
      filters: [{ name: 'Documents', extensions: [...ATTACH_EXTS] }],
    })
    if (res.canceled || res.filePaths.length === 0) return null
    const accepted: AttachmentMeta[] = []
    const rejected: string[] = []
    for (const p of res.filePaths) {
      const name = basename(p)
      const ext = extname(p).replace('.', '').toLowerCase()
      if (!ATTACH_EXTS.has(ext)) { rejected.push(`${name}: unsupported type`); continue }
      try {
        const st = statSync(p)
        if (st.size > ATTACH_MAX) { rejected.push(`${name}: too large`); continue }
        accepted.push({ path: p, name, ext, sizeBytes: st.size })
      } catch { rejected.push(`${name}: unreadable`) }
    }
    return { accepted, rejected }
  })

  ipcMain.handle('files:read', async (_e, path: string) => {
    try {
      const parsed = await parseFileToText(String(path))
      if (parsed.kind === 'image') return { ok: true, kind: 'image' as const }
      if (parsed.kind === 'unsupported') return { ok: false, kind: 'unsupported' as const, error: 'Unsupported file' }
      return { ok: true, kind: 'text' as const, text: parsed.text ?? '' }
    } catch (err) {
      return { ok: false, kind: 'text' as const, error: String(err) }
    }
  })

  ipcMain.handle('files:read-image', async (_e, path: string) => {
    try {
      const ext = extname(String(path)).replace('.', '').toLowerCase()
      const mime = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : ext === 'webp' ? 'image/webp' : 'image/jpeg'
      const bytes = readFileSync(String(path))
      return { ok: true, base64: bytes.toString('base64'), mime }
    } catch (err) {
      return { ok: false, error: String(err) }
    }
  })

  // no-op auth handlers kept for parity with the shared AI settings UI
  ipcMain.handle('ai:gsk-status', () => ({ loggedIn: false }))
  ipcMain.handle('ai:gsk-login', () => {
    void shell.openExternal('https://llm.proxsis.com')
  })
}
