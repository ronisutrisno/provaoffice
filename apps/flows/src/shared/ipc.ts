/** Shared IPC channels + types for the flows (AI flowchart) app. */
import type {
  AgentMessage,
  AgentToolCall,
  AgentToolDef,
  IpcStreamChunk,
  IpcStreamStart,
} from '@prova/agent-core'
import type { AiSettings } from '@prova/ai-provider'

export type { AgentMessage, AgentToolCall, AgentToolDef, AiSettings }

export type NodeType = 'start' | 'end' | 'process' | 'decision' | 'io'

export interface FlowNode {
  id: string
  type: NodeType
  label: string
  x: number
  y: number
  lane?: string | undefined
}

export interface FlowEdge {
  id: string
  from: string
  to: string
  label?: string
}

export interface FlowLane {
  id: string
  name: string
}

export interface FlowDocument {
  version: 1
  title: string
  lanes?: FlowLane[] | undefined
  nodes: FlowNode[]
  edges: FlowEdge[]
}

export function emptyFlow(): FlowDocument {
  return { version: 1, title: 'Untitled Flow', nodes: [], edges: [] }
}

/** Node size per type (px) - used by both the renderer layout and the AI auto-layout. */
export const NODE_SIZE: Record<NodeType, { w: number; h: number }> = {
  start: { w: 120, h: 56 },
  end: { w: 120, h: 56 },
  process: { w: 160, h: 72 },
  decision: { w: 160, h: 96 },
  io: { w: 160, h: 72 },
}

export const ATTACHMENT_IMAGE_EXTS = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp'])

export interface AttachmentMeta {
  path: string
  name: string
  ext: string
  sizeBytes: number
}

export interface AttachmentAddResult {
  accepted: AttachmentMeta[]
  rejected: string[]
}

export interface AttachmentReadResult {
  ok: boolean
  kind: 'text' | 'image' | 'unsupported'
  text?: string
  error?: string
}

export interface AttachmentImageResult {
  ok: boolean
  base64?: string
  mime?: string
  error?: string
}

export interface WebSearchResponse {
  results: Array<{ title: string; url: string; snippet: string }>
  answer?: string | undefined
  method: string
  error?: string | undefined
}

export interface ImageSearchResponse {
  images: Array<{
    title: string
    imageUrl: string
    sourceUrl: string
    source: string
    width?: number | undefined
    height?: number | undefined
  }>
  method: string
  error?: string | undefined
}

export interface FlowsApi {
  getLanguage(): Promise<string>
  getTheme(): Promise<'system' | 'light' | 'dark'>
  onThemeChanged(cb: (theme: 'system' | 'light' | 'dark') => void): () => void
  newFlow(): Promise<void>
  openFlow(): Promise<{ path: string; doc: FlowDocument } | null>
  saveFlow(doc: FlowDocument): Promise<{ path: string } | null>
  saveFlowAs(doc: FlowDocument): Promise<{ path: string } | null>
  exportPng(dataUrl: string): Promise<{ path: string } | null>
  exportSvg(svg: string): Promise<{ path: string } | null>
  consumePendingOpen(): Promise<{ path: string; doc: FlowDocument } | null>
  onMenu(cb: (cmd: string) => void): () => void
  aiStream(request: IpcStreamStart<AiSettings>): Promise<void>
  onAiStream(cb: (chunk: IpcStreamChunk) => void): () => void
  aiStreamCancel(requestId: string): Promise<void>
  getAiSettings(): Promise<AiSettings>
  setAiSettings(settings: AiSettings): Promise<void>
  pickAttachments(): Promise<AttachmentAddResult | null>
  readAttachment(path: string): Promise<AttachmentReadResult>
  readAttachmentImage(path: string): Promise<AttachmentImageResult>
  webSearch(query: string, maxResults?: number): Promise<WebSearchResponse>
  imageSearch(query: string, maxResults?: number): Promise<ImageSearchResponse>
}

declare global {
  interface Window {
    flowsApi: FlowsApi
  }
}