import type { AgentToolDef } from '../../shared/ipc'
import {
  DEFAULT_THEME,
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
      'Create a new presentation. Generates a cover slide and initializes the deck. Always call this first.',
    inputSchema: {
      type: 'object',
      properties: {
        filename: { type: 'string', description: 'Presentation title (used as deck name)' },
        title: { type: 'string', description: 'Cover slide title' },
        subtitle: { type: 'string', description: 'Cover slide subtitle (optional)' },
        theme: {
          type: 'object',
          description: 'Color theme override: { primary, secondary, accent, background } as hex colors',
          properties: {
            primary: { type: 'string' },
            secondary: { type: 'string' },
            accent: { type: 'string' },
            background: { type: 'string' },
          },
        },
      },
      required: ['filename', 'title'],
    },
  },
  {
    name: 'add_slide_with_content',
    description:
      'Add a content slide to the presentation. Layouts: title_content (bullets+image), two_column, cards (3 cards with title+desc), rows (list with markers), stats (big numbers), section_header, closing, agenda, blank. Content format depends on layout: strings for bullets, {title,desc}[] for cards/rows, {big,desc}[] for stats. IMPORTANT: always provide image_query with English keywords for EVERY content slide so a relevant photo is placed on the slide.',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Slide title' },
        content: {
          type: 'array',
          description: 'Content items. Format: ["bullet1","bullet2"] for bullets, [{"title":"X","desc":"Y"}] for cards/rows, [{"big":"$30T","desc":"value"}] for stats',
          items: { type: 'string' },
        },
        layout: {
          type: 'string',
          description: 'Layout type: title_content (default), two_column, cards, rows, stats, section_header, closing, agenda, blank',
        },
        eyebrow: { type: 'string', description: 'Small label above title (e.g. "01 · PENGANTAR")' },
        intro: { type: 'string', description: 'Intro paragraph below title' },
        subtitle: { type: 'string', description: 'Subtitle for closing layout' },
        image_query: { type: 'string', description: 'REQUIRED. English Pixabay search keywords for the slide photo, e.g. "solar panels green energy" or "corporate meeting boardroom". Always provide this for every content slide.' },
      },
      required: ['title', 'content', 'layout', 'image_query'],
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

const PIXABAY_KEY = '55586367-a4b8c80ba306e0b4fb4afca94'

async function fetchPixabayImage(query: string): Promise<string | undefined> {
  try {
    const url = `https://pixabay.com/api/?key=${PIXABAY_KEY}&q=${encodeURIComponent(query)}&image_type=photo&per_page=3&safesearch=true`
    const resp = await fetch(url)
    if (!resp.ok) return undefined
    const data = await resp.json()
    const hit = data.hits?.[0]
    return hit?.webformatURL ?? hit?.largeImageURL ?? undefined
  } catch {
    return undefined
  }
}

/** Fallback English Pixabay keywords derived from the slide title — used only when the
 * model omits image_query AND the layout benefits from a photo (cover/section/title_content). */
function deriveImageQuery(title: string, layout: string): string | undefined {
  if (layout !== 'title' && layout !== 'section_header' && layout !== 'title_content') return undefined
  const t = title.toLowerCase()
  const map: Array<[RegExp, string]> = [
    [/esg|environment|sustainab|green|climate|karbon|lingkungan/, 'green energy sustainability'],
    [/social|human|people|community|diversity|sosial/, 'people community teamwork'],
    [/governance|corporate|company|management|tata kelola/, 'corporate meeting boardroom'],
    [/market|invest|finance|stock|pasar|investasi/, 'stock market finance'],
    [/indonesia|jakarta|nusantara/, 'indonesia jakarta city'],
    [/energy|solar|wind|renewable|energi/, 'solar panels renewable energy'],
    [/chart|data|statistik|grafik/, 'data analytics chart'],
    [/tech|digital|ai|teknologi/, 'technology digital innovation'],
    [/health|medis|kesehatan/, 'healthcare medical'],
    [/education|learning|pendidikan/, 'education learning students'],
    [/tren|trend|2026|2025|proyeksi/, 'business growth trend'],
  ]
  for (const [re, kw] of map) if (re.test(t)) return kw
  return 'business presentation'
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
      currentDeckName = String(input.filename ?? title).trim()
      if (input.theme && typeof input.theme === 'object') {
        const th = input.theme as Record<string, string>
        currentTheme = { ...DEFAULT_THEME }
        if (th.primary) currentTheme.primary = th.primary
        if (th.secondary) currentTheme.secondary = th.secondary
        if (th.accent) currentTheme.accent = th.accent
        if (th.background) currentTheme.background = th.background
      }
      const coverQuery = deriveImageQuery(title, 'title')
      const coverImage = coverQuery ? await fetchPixabayImage(coverQuery) : undefined
      currentSlides = [{ title, subtitle, content: [], layout: 'title', imageUrl: coverImage, theme: { ...currentTheme } }]
      return {
        output: `Presentation "${currentDeckName}" created with cover slide. Title: "${title}". Use add_slide_with_content to add more slides.`,
        mutated: true,
        summary: `Created presentation: ${title}`,
      }
    }

    case 'add_slide_with_content': {
      const title = String(input.title ?? '').trim()
      if (!title) return { output: 'title is required', summary: 'Error: title required' }
      const layout = (String(input.layout ?? 'title_content').trim()) as SlideContent['layout']
      const eyebrow = String(input.eyebrow ?? '').trim() || undefined
      const intro = String(input.intro ?? '').trim() || undefined
      const subtitle = String(input.subtitle ?? '').trim() || undefined
      const content = Array.isArray(input.content) ? input.content : []
      const imageQuery = String(input.image_query ?? '').trim() || undefined

      let parsedContent: SlideContent['content']
      if (layout === 'cards' || layout === 'rows') {
        parsedContent = content.map(c => {
          if (typeof c === 'object' && c !== null && 'title' in (c as any)) return c as { title: string; desc?: string }
          return { title: String(c), desc: '' }
        })
      } else if (layout === 'stats') {
        parsedContent = content.map(c => {
          if (typeof c === 'object' && c !== null && 'big' in (c as any)) return c as { big: string; desc: string }
          return { big: String(c), desc: '' }
        })
      } else {
        parsedContent = content.map(c => String(c))
      }

      let imageUrl: string | undefined
      if (imageQuery) imageUrl = await fetchPixabayImage(imageQuery)
      else {
        const fallback = deriveImageQuery(title, layout)
        if (fallback) imageUrl = await fetchPixabayImage(fallback)
      }

      const slide: SlideContent = {
        title,
        subtitle,
        eyebrow,
        intro,
        content: parsedContent,
        layout,
        imageUrl,
        theme: { ...currentTheme },
      }

      const slideNum = currentSlides.length + 1
      currentSlides.push(slide)
      return {
        output: `Slide ${slideNum} added: "${title}" (layout: ${layout}). Total slides: ${currentSlides.length}.${imageUrl ? ` Image: ${imageUrl}` : ''}`,
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
