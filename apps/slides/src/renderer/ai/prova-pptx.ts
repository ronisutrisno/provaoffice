import type { AgentToolDef } from '../../shared/ipc'
import {
  DEFAULT_THEME,
  THEME_PRESETS,
  redirectLayout,
  type BusinessCasePayload,
  type SlideContent,
  type SlideTheme,
} from './slide-templates'

/** State for current presentation being built */
let currentTheme: SlideTheme = { ...DEFAULT_THEME }
let currentSlides: SlideContent[] = []
let currentDeckName = ''

export function resetPresentation(): void {
  currentTheme = { ...DEFAULT_THEME }
  currentSlides = []
  currentDeckName = ''
}

export function getCurrentDeckName(): string {
  return currentDeckName
}

const PROVA_JSON_PREFIX = 'prova-json:'

function encodeSlideMarker(slide: SlideContent): string {
  return PROVA_JSON_PREFIX + btoa(unescape(encodeURIComponent(JSON.stringify(slide))))
}

export const PROVA_PPTX_TOOLS: AgentToolDef[] = [
  {
    name: 'create_presentation',
    description:
      'Create a new presentation. Generates a cover slide (maroon/brand background + optional photo panel) and initializes the deck. Always call this first. Pick a theme_preset by name — do not invent colors.',
    inputSchema: {
      type: 'object',
      properties: {
        filename: { type: 'string', description: 'Presentation title (used as deck name)' },
        title: { type: 'string', description: 'Cover slide title' },
        subtitle: { type: 'string', description: 'Cover slide subtitle (optional)' },
        eyebrow: { type: 'string', description: 'Small label above cover title (optional, default "PRESENTASI")' },
        image_query: { type: 'string', description: 'Optional English Pixabay keywords for the cover photo panel (right side ~38%). Cover works beautifully without a photo too.' },
        theme_preset: {
          type: 'string',
          enum: ['maroon', 'navy', 'green'],
          description:
            'Color set for the whole deck (house style is identical, only the hue changes). maroon = deep maroon/brand default (business, government, finance); navy = dark navy blue (technology, corporate, data); green = forest green (health, sustainability, agriculture, education). Default maroon.',
        },
      },
      required: ['filename', 'title'],
    },
  },
  {
    name: 'add_slide_with_content',
    description:
      'Add a content slide. House-style layouts (all native & editable, Georgia font): numbered_list (default — rounded rows with big 01/02 numbers + accent bar; content = ["text"] or [{title,desc}], max 5), cards (2-3 uniform white cards with colored band; content=[{title,desc}] max 3), stats (up to 4 big-number cards; content=[{big,desc}]), callout (quote banner; content=["quote text"], intro=attribution), agenda (numbered table of contents; content=["item",...]), business_case (RICHEST layout — banner KEY MESSAGE + contrast cards + optional table; content=[{...payload dict}] with fields variant full|solution|metrics, badge, key_message, background, business_problem, solution[], table_title, table{headers,rows[]}, benefits[], metrics[], challenge, impact), section_header (chapter divider — call between chapters), closing, blank. The old layouts title_content/two_column/rows/timeline/comparison/quote/big_number are accepted but redirected to the new equivalents. Images are NOT required: they are added only when the slide is sparse (<=3 short rows) and there is room — set image_query only for such slides; never force an image. Content density: every slide must carry >= 3 facts; desc items must be informative 10-25 words; a title promising N items must list all N in the content.',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Slide title' },
        content: {
          type: 'array',
          description:
            'Content items. ["bullet"] or [{title,desc}] for numbered_list/cards; [{big,desc}] for stats; ["quote"] for callout; for business_case pass ONE dict payload: [{variant:"full",badge:"...",key_message:"...",background:"...",business_problem:"...",solution:["p1","p2"],table_title:"...",table:{headers:[...],rows:[[...]]}}]',
          items: { type: 'string' },
        },
        layout: {
          type: 'string',
          description:
            'numbered_list (default), cards, stats, callout, agenda, business_case, section_header, closing, blank',
        },
        eyebrow: { type: 'string', description: 'Small caps label above title (e.g. "01 · LATAR BELAKANG")' },
        intro: { type: 'string', description: 'One-line framing sentence — rendered as a KEY MESSAGE banner' },
        subtitle: { type: 'string', description: 'Subtitle for closing layout' },
        image_query: { type: 'string', description: 'Optional English Pixabay keywords. ONLY for sparse slides (<=3 short rows). Leave empty when content is dense — the system auto-adds a photo to sparse slides without one.' },
        image_side: { type: 'string', enum: ['left', 'right'], description: 'Photo side; alternate across slides. Default right.' },
      },
      required: ['title'],
    },
  },
  {
    name: 'add_closing',
    description: 'Add a closing/thank you slide.',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Closing title (default: "Terima Kasih")' },
        subtitle: { type: 'string', description: 'Contact info or tagline' },
      },
      required: [],
    },
  },
  {
    name: 'list_slides',
    description: 'List all current slides in the deck with their titles.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
]

export interface ProvaToolResult {
  output: string
  mutated?: boolean
  summary: string
}

async function fetchPixabayImage(query: string): Promise<string | undefined> {
  try {
    // search runs in the main process (Pixabay primary) — the renderer CSP blocks direct fetch.
    // Pick from the top 3 hits (seeded per call) instead of always #1: Pixabay ranks the same
    // stock photo first for generic queries, which made every deck wear the same cover photo.
    const r = await window.slidesApi.imageSearch(query, 3)
    const imgs = r.images
    if (imgs.length === 0) return undefined
    const pick = Math.floor(Math.random() * imgs.length)
    return imgs[pick]?.imageUrl
  } catch {
    return undefined
  }
}

/**
 * Topic-based theme fallback when the model omits/misspells theme_preset.
 * Only three house-style color sets exist now: maroon / navy / green.
 */
function deriveThemePreset(title: string): string {
  const t = title.toLowerCase()
  if (/health|medis|kesehatan|dokter|rs\b|rumah sakit|patient|pasien|climate|energy|sustainab|green|lingkungan|karbon|esg|pertanian|education|pendidikan|sekolah/.test(t)) return 'green'
  if (/tech|digital|ai\b|data|software|startup|innovation|teknologi|market|invest|finance|stock|pasar|investasi|project|manajemen|management|pmi|operations|audit|keuangan|finance/.test(t)) return 'navy'
  return 'maroon'
}

/** Auto Pixabay keywords for SPARSE slides only (<=3 short rows) — never forced on dense slides. */
function deriveImageQuery(title: string): string {
  const t = title.toLowerCase()
  const map: Array<[RegExp, string]> = [
    [/esg|environment|sustainab|green|climate|karbon|lingkungan/, 'green energy sustainability'],
    [/social|human|people|community|diversity|sosial/, 'people community teamwork'],
    [/governance|corporate|company|management|tata kelola/, 'corporate meeting boardroom'],
    [/market|invest|finance|stock|pasar|investasi/, 'stock market finance'],
    [/indonesia|jakarta|nusantara/, 'indonesia jakarta city skyline'],
    [/energy|solar|wind|renewable|energi/, 'solar panels renewable energy'],
    [/chart|data|statistik|grafik|kpi|angka|metric/, 'data analytics dashboard'],
    [/tech|digital|ai|teknologi/, 'technology digital innovation'],
    [/health|medis|kesehatan|keracunan|pangan|food/, 'food safety hygiene kitchen'],
    [/education|learning|pendidikan|sekolah/, 'education learning students classroom'],
    [/tren|trend|2026|2025|proyeksi|roadmap/, 'business growth strategy planning'],
  ]
  for (const [re, kw] of map) if (re.test(t)) return kw
  return 'professional business meeting'
}

/** Count renderable facts in content — dense slides must not get an image. */
function contentVolume(content: SlideContent['content']): { n: number; chars: number } {
  const items = Array.isArray(content) ? content : []
  let chars = 0
  for (const c of items) {
    if (typeof c === 'string') chars += c.length
    else if (c && typeof c === 'object') {
      const o = c as Record<string, unknown>
      chars += String(o.title ?? o.big ?? '').length + String(o.desc ?? '').length
    }
  }
  return { n: items.length, chars }
}

export async function executeProvaTool(
  name: string,
  input: Record<string, unknown>,
): Promise<ProvaToolResult> {
  switch (name) {
    case 'create_presentation': {
      const title = String(input.title ?? '').trim()
      if (!title) return { output: 'title is required', summary: 'Error: title required' }
      const subtitle = String(input.subtitle ?? '').trim()
      const eyebrow = String(input.eyebrow ?? '').trim() || undefined
      currentDeckName = String(input.filename ?? title).trim()
      const presetName = String(input.theme_preset ?? '').trim().toLowerCase()
      currentTheme = { ...(THEME_PRESETS[presetName] ?? THEME_PRESETS[deriveThemePreset(title)] ?? DEFAULT_THEME) }
      const coverQuery = String(input.image_query ?? '').trim()
      const coverImage = coverQuery ? await fetchPixabayImage(coverQuery) : undefined
      currentSlides = [{ title, subtitle, eyebrow, content: [], layout: 'title', imageUrl: coverImage, theme: { ...currentTheme } }]
      return {
        output: `Presentation "${currentDeckName}" created with cover slide (theme: ${presetName || deriveThemePreset(title)}). Use add_slide_with_content to add more slides.`,
        mutated: true,
        summary: `Created presentation: ${title}`,
      }
    }

    case 'add_slide_with_content': {
      const title = String(input.title ?? '').trim()
      if (!title) return { output: 'title is required', summary: 'Error: title required' }
      const rawLayout = String(input.layout ?? 'numbered_list').trim()
      const layout = redirectLayout(rawLayout)
      const eyebrow = String(input.eyebrow ?? '').trim() || undefined
      const intro = String(input.intro ?? '').trim() || undefined
      const subtitle = String(input.subtitle ?? '').trim() || undefined
      const content = Array.isArray(input.content) ? input.content : []
      let imageQuery = String(input.image_query ?? '').trim() || undefined
      const imageSide = input.image_side === 'left' ? 'left' : 'right'

      // business_case: content = satu dict payload (toleran: [dict], dict, string JSON)
      let bc: BusinessCasePayload | undefined
      if (layout === 'business_case') {
        let rc: unknown = content
        if (Array.isArray(rc) && rc.length) rc = rc[0]
        if (typeof rc === 'string' && rc.trim().startsWith('{')) {
          try { rc = JSON.parse(rc) } catch { /* keep string */ }
        }
        if (!rc || typeof rc !== 'object' || Array.isArray(rc)) {
          return {
            output:
              'Error: layout business_case membutuhkan content satu dict: [{variant:"full|solution|metrics",badge,key_message,background,business_problem,solution[],table_title,table:{headers,rows},benefits[],metrics[],challenge,impact}] — minimal key_message atau background.',
            summary: 'Error: business_case payload salah',
          }
        }
        bc = rc as BusinessCasePayload
      }

      let parsedContent: SlideContent['content']
      if (layout === 'stats') {
        parsedContent = content.map((c) => {
          if (c && typeof c === 'object' && 'big' in (c as Record<string, unknown>)) return c as { big: string; desc: string }
          const o = c as Record<string, unknown>
          if (o && typeof o === 'object' && (o.title || o.desc)) return { big: String(o.title ?? ''), desc: String(o.desc ?? '') }
          return { big: String(c), desc: '' }
        })
      } else if (layout === 'cards' || layout === 'numbered_list') {
        parsedContent = content.map((c) => {
          if (c && typeof c === 'object' && ('title' in (c as Record<string, unknown>) || 'desc' in (c as Record<string, unknown>))) return c as { title: string; desc?: string }
          return String(c)
        }) as SlideContent['content']
      } else {
        parsedContent = content.map((c) => String(c))
      }

      // VALIDASI: layout konten harus punya sesuatu yang bisa dirender —
      // slide hanya-judul+hanya-intro DITOLAK (guardrail ported from fastapi).
      if (layout !== 'section_header' && layout !== 'closing' && layout !== 'blank' && layout !== 'business_case') {
        const { n } = contentVolume(parsedContent)
        if (n === 0) {
          return {
            output: `Error: layout '${layout}' tidak punya konten — slide hanya judul/inti DITOLAK. Kirim content nyata: ["poin",...] atau [{title,desc}] atau [{big,desc}].`,
            summary: 'Error: konten kosong',
          }
        }
      }
      if (layout === 'business_case') {
        const hasAny = bc && (bc.key_message || bc.background || bc.solution || bc.benefits || bc.metrics || bc.table || bc.challenge || bc.impact)
        if (!hasAny) return { output: 'Error: business_case payload kosong — isi minimal key_message atau background.', summary: 'Error: business_case kosong' }
      }

      // Gambar: HANYA diminta eksplisit ATAU slide sepi (<=3 baris pendek) → auto-isi.
      let imageUrl: string | undefined
      const wantsImage = layout !== 'stats' && layout !== 'callout' && layout !== 'business_case' && layout !== 'section_header'
      if (wantsImage) {
        const vol = contentVolume(parsedContent)
        const sparse = vol.n <= 3 && vol.chars < 260
        const query = imageQuery ?? (sparse ? deriveImageQuery(title) : undefined)
        if (query) imageUrl = await fetchPixabayImage(query)
      }

      const slide: SlideContent = {
        title,
        subtitle,
        eyebrow,
        intro,
        content: parsedContent,
        layout,
        imageUrl,
        imageSide,
        theme: { ...currentTheme },
        ...(bc ? { bc } : {}),
      }

      const slideNum = currentSlides.length + 1
      currentSlides.push(slide)
      return {
        output: `Slide ${slideNum} added: "${title}" (layout: ${layout}${rawLayout !== layout ? `, redirected from ${rawLayout}` : ''}). Total slides: ${currentSlides.length}.${imageUrl ? ` Image: ${imageUrl}` : ''}`,
        mutated: true,
        summary: `Added slide: ${title}`,
      }
    }

    case 'add_closing': {
      const title = String(input.title ?? 'Terima Kasih').trim() || 'Terima Kasih'
      const subtitle = String(input.subtitle ?? '').trim()
      currentSlides.push({ title, subtitle, content: [], layout: 'closing', theme: { ...currentTheme } })
      return {
        output: `Closing slide added. Total slides: ${currentSlides.length}.`,
        mutated: true,
        summary: `Added closing slide`,
      }
    }

    case 'list_slides': {
      if (currentSlides.length === 0) {
        return { output: 'No slides yet. Call create_presentation first.', summary: 'No slides' }
      }
      return {
        output: `Deck "${currentDeckName}": ${currentSlides.length} slide(s).`,
        summary: `Listed ${currentSlides.length} slides`,
      }
    }

    case 'delete_slide': {
      const idx = Number(input.slideIndex ?? -1)
      if (idx < 0 || idx >= currentSlides.length) {
        return { output: `Invalid index. Range: 0-${currentSlides.length - 1}`, summary: 'Invalid index' }
      }
      currentSlides.splice(idx, 1)
      return {
        output: `Slide ${idx + 1} removed. Total slides: ${currentSlides.length}.`,
        mutated: true,
        summary: `Removed slide ${idx + 1}`,
      }
    }

    default:
      return { output: `Unknown tool: ${name}`, summary: `Unknown tool: ${name}` }
  }
}

/** Get all accumulated slides as prova-json markers for generateFromHtml */
export function getSlidesHtml(): string[] {
  return currentSlides.map(encodeSlideMarker)
}
