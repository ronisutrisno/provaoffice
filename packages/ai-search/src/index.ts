/**
 * Search utilities (main process) — Ollama Web Search API first (https://ollama.com/api/web_search),
 * then Serper Google API, with DuckDuckGo as the last resort. The Serper/DuckDuckGo logic mirrors an
 * earlier web_search / google_image_search implementation. Runs in the main process
 * (Node fetch / child process) to avoid renderer CORS; the Serper key reuses SERPER_API_KEY.
 * For gsk auth see ./gsk.ts (`gsk login` or GSK_API_KEY).
 */

import {
  COPYRIGHT_HOSTS,
  asRecord,
  safeHost,
  type ImageSearchResult,
  type WebSearchResult,
} from './shared'
import { gskImageSearch, gskWebSearch, hasGskAuth } from './gsk'

export type { ImageSearchResult, WebSearchResult } from './shared'
export * from './gsk'
export * from './genoffice-auth'

const SERPER_KEY = () => process.env.SERPER_API_KEY ?? ''
const OLLAMA_API_KEY = '71a9d89a49ec49faa633d53d9b9dc650._sSayBGVLMFtcMXxAL70gSa3'
const OLLAMA_ENDPOINT = 'https://ollama.com/api/web_search'
const OLLAMA_MAX_RESULTS = 10

// ── Web search ──────────────────────────────────────────────────────

export async function webSearch(
  query: string,
  maxResults = 6,
): Promise<{
  results: WebSearchResult[]
  answer?: string
  method: string
}> {
  // 1) Ollama Web Search (primary — same stack as proxsis-search-mcp1)
  try {
    const r = await ollamaWebSearch(query, maxResults)
    if (r.results.length) return { ...r, method: 'ollama' }
  } catch {
    /* fall back to Serper/DuckDuckGo */
  }
  // 2) gsk (PROVA-AI CLI)
  if (hasGskAuth()) {
    try {
      const r = await gskWebSearch(query, maxResults)
      if (r.results.length) return { ...r, method: 'gsk' }
    } catch {
      /* fall back to Serper/DuckDuckGo */
    }
  }
  const key = SERPER_KEY()
  if (key) {
    try {
      const resp = await fetchWithTimeout('https://google.serper.dev/search', {
        method: 'POST',
        headers: { 'X-API-KEY': key, 'Content-Type': 'application/json' },
        body: JSON.stringify({ q: query, num: maxResults, gl: 'us', hl: 'en' }),
      })
      if (resp.ok) {
        const data = asRecord(await resp.json())
        const organic: unknown[] = Array.isArray(data.organic) ? data.organic : []
        const results: WebSearchResult[] = organic.slice(0, maxResults).map((item) => {
          const o = asRecord(item)
          return {
            title: String(o.title ?? ''),
            url: String(o.link ?? ''),
            snippet: String(o.snippet ?? ''),
          }
        })
        const answerBox = asRecord(data.answerBox)
        const answerRaw =
          answerBox.answer || answerBox.snippet || asRecord(data.knowledgeGraph).description
        const answer = typeof answerRaw === 'string' && answerRaw ? answerRaw : undefined
        if (results.length) {
          return answer !== undefined
            ? { results, answer, method: 'serper' }
            : { results, method: 'serper' }
        }
      }
    } catch {
      /* fall back to DuckDuckGo */
    }
  }
  return { ...(await duckWebSearch(query, maxResults)), method: 'duckduckgo' }
}

/** Ollama Web Search — https://ollama.com/api/web_search (same stack as proxsis-search-mcp1).
 * Only `query` + `max_results` (max 10) are accepted upstream; language/region are
 * encoded into the query text as an international-English bias. */
async function ollamaWebSearch(
  query: string,
  maxResults: number,
): Promise<{ results: WebSearchResult[] }> {
  const q = (query || '').trim().slice(0, 400)
  if (!q) return { results: [] }
  const limit = Math.min(Math.max(1, maxResults), OLLAMA_MAX_RESULTS)
  const internationalBias = 'international English perspective global'
  const payload = { query: `${q} ${internationalBias}`.slice(0, 600), max_results: limit }
  const resp = await fetchWithTimeout(OLLAMA_ENDPOINT, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${OLLAMA_API_KEY}`,
      'Content-Type': 'application/json',
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    },
    body: JSON.stringify(payload),
    timeoutMs: 35000,
  })
  if (!resp.ok) return { results: [] }
  const text = await resp.text()
  if (!text.trim()) return { results: [] }
  const data = asRecord(JSON.parse(text))
  const raw: unknown[] = Array.isArray(data.results) ? data.results : []
  const results: WebSearchResult[] = []
  for (const item of raw) {
    const o = asRecord(item)
    const title = String(o.title ?? 'Info').trim()
    const url = String(o.url ?? '').trim()
    const content = String(o.content ?? o.snippet ?? '').trim()
    if (!url || (!title && !content)) continue
    results.push({ title, url, snippet: content.slice(0, 1400) })
    if (results.length >= limit) break
  }
  return { results }
}

// ── Image search ────────────────────────────────────────────────────

export async function imageSearch(
  query: string,
  maxResults = 8,
): Promise<{
  images: ImageSearchResult[]
  method: string
}> {
  if (hasGskAuth()) {
    try {
      const images = await gskImageSearch(query, maxResults)
      if (images.length) return { images, method: 'gsk' }
    } catch {
      /* fall back to Serper/DuckDuckGo */
    }
  }
  const key = SERPER_KEY()
  if (key) {
    try {
      const resp = await fetchWithTimeout('https://google.serper.dev/images', {
        method: 'POST',
        headers: { 'X-API-KEY': key, 'Content-Type': 'application/json' },
        body: JSON.stringify({ q: query, num: Math.min(maxResults, 10), gl: 'us', hl: 'en' }),
      })
      if (resp.ok) {
        const data = asRecord(await resp.json())
        const raw: unknown[] = Array.isArray(data.images) ? data.images : []
        const images: ImageSearchResult[] = []
        for (const item of raw) {
          const img = asRecord(item)
          const imageUrl = String(img.imageUrl ?? img.original ?? '')
          if (!imageUrl) continue
          if (COPYRIGHT_HOSTS.some((d) => imageUrl.toLowerCase().includes(d))) continue
          const entry: ImageSearchResult = {
            title: String(img.title ?? ''),
            imageUrl,
            sourceUrl: String(img.link ?? ''),
            source: String(img.source ?? safeHost(img.link)),
          }
          if (typeof img.imageWidth === 'number') entry.width = img.imageWidth
          if (typeof img.imageHeight === 'number') entry.height = img.imageHeight
          images.push(entry)
          if (images.length >= maxResults) break
        }
        if (images.length) return { images, method: 'serper' }
      }
    } catch {
      /* fall back to DuckDuckGo */
    }
  }
  return { images: await duckImageSearch(query, maxResults), method: 'duckduckgo' }
}

// ── DuckDuckGo fallback (no key / quota exhausted) ──────────────────

async function duckWebSearch(
  query: string,
  maxResults: number,
): Promise<{ results: WebSearchResult[] }> {
  try {
    // DuckDuckGo HTML endpoint (lightweight, no key needed)
    const resp = await fetchWithTimeout(
      `https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}`,
      { headers: { 'User-Agent': 'Mozilla/5.0' } },
    )
    const html = await resp.text()
    const results: WebSearchResult[] = []
    const re = /<a[^>]*class="result__a"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(html)) !== null && results.length < maxResults) {
      const url = decodeDuckUrl(m[1]!)
      const title = stripTags(m[2]!)
      if (url && title) results.push({ title, url, snippet: '' })
    }
    return { results }
  } catch {
    return { results: [] }
  }
}

async function duckImageSearch(query: string, maxResults: number): Promise<ImageSearchResult[]> {
  try {
    // DuckDuckGo i.js needs a vqd token, so it takes two steps
    const tokenResp = await fetchWithTimeout(
      `https://duckduckgo.com/?q=${encodeURIComponent(query)}`,
      { headers: { 'User-Agent': 'Mozilla/5.0' } },
    )
    const tokenHtml = await tokenResp.text()
    const vqd = /vqd=["']?([\d-]+)["']?/.exec(tokenHtml)?.[1]
    if (!vqd) return []
    const resp = await fetchWithTimeout(
      `https://duckduckgo.com/i.js?l=us-en&o=json&q=${encodeURIComponent(query)}&vqd=${vqd}`,
      { headers: { 'User-Agent': 'Mozilla/5.0', Referer: 'https://duckduckgo.com/' } },
    )
    const data = asRecord(await resp.json())
    const list: unknown[] = Array.isArray(data.results) ? data.results : []
    const out: ImageSearchResult[] = []
    for (const item of list.slice(0, maxResults)) {
      const img = asRecord(item)
      const imageUrl = String(img.image ?? '')
      if (!imageUrl || COPYRIGHT_HOSTS.some((d) => imageUrl.toLowerCase().includes(d))) continue
      const entry: ImageSearchResult = {
        title: String(img.title ?? ''),
        imageUrl,
        sourceUrl: String(img.url ?? ''),
        source: safeHost(img.url),
      }
      if (typeof img.width === 'number') entry.width = img.width
      if (typeof img.height === 'number') entry.height = img.height
      out.push(entry)
    }
    return out
  } catch {
    return []
  }
}

// ── utils ───────────────────────────────────────────────────────────

async function fetchWithTimeout(
  url: string,
  init: RequestInit & { timeoutMs?: number } = {},
): Promise<Response> {
  const controller = new AbortController()
  const t = setTimeout(() => controller.abort(), init.timeoutMs ?? 15000)
  try {
    return await fetch(url, { ...init, signal: controller.signal })
  } finally {
    clearTimeout(t)
  }
}

function stripTags(s: string): string {
  return s
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&#x27;/g, "'")
    .trim()
}

function decodeDuckUrl(href: string): string {
  // DuckDuckGo result links are often /l/?uddg=<encoded>
  const m = /[?&]uddg=([^&]+)/.exec(href)
  if (m) return decodeURIComponent(m[1]!)
  return href.startsWith('http') ? href : ''
}
