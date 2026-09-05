import type { AgentToolDef } from '../../shared/ipc'
import {
  DEFAULT_THEME,
  THEME_PRESETS,
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
      'Create a new presentation. Generates a cover slide and initializes the deck. Always call this first. Pick a theme_preset by name (recommended) — do not invent colors. Custom theme colors are allowed only when the user explicitly asks for specific colors.',
    inputSchema: {
      type: 'object',
      properties: {
        filename: { type: 'string', description: 'Presentation title (used as deck name)' },
        title: { type: 'string', description: 'Cover slide title' },
        subtitle: { type: 'string', description: 'Cover slide subtitle (optional)' },
        image_query: { type: 'string', description: 'English Pixabay keywords for the COVER photo (recommended, e.g. "medical students lecture hall classroom"). Pick something that reflects the deck topic — do not rely on the generic fallback.' },
        theme_preset: {
          type: 'string',
          enum: ['corporate', 'ocean', 'forest', 'sunset', 'slate'],
          description: 'Curated color palette for the whole deck. corporate = navy/blue (default, business), ocean = deep blue/amber, forest = green/gold, sunset = plum/terracotta, slate = gray-green/coral. Pick one that fits the topic and use it for the whole deck.',
        },
        theme: {
          type: 'object',
          description: 'Custom color override (only when the user explicitly requests specific colors): { primary, secondary, accent, background } as hex colors. Keep background light (or a dark brand color, not black) so text stays readable.',
          properties: {
            primary: { type: 'string' },
            secondary: { type: 'string' },
            accent: { type: 'string' },
            background: { type: 'string' },
          },
        },
      },
      required: ['filename', 'title', 'theme_preset'],
    },
  },
  {
    name: 'add_slide_with_content',
    description:
      'Add a content slide to the presentation. Layouts: title_content (bullets+image), two_column (2 text columns), cards (3 cards), rows (list with markers), stats (up to 4 big numbers), timeline (horizontal steps, alternating above/below), quote (pull quote; content=[quote text], intro=attribution), big_number (one hero figure; content=[number], intro=caption), comparison (2 panels e.g. before/after), section_header, closing, agenda, blank. Content format: strings for bullets, {title,desc}[] for cards/rows/timeline/comparison, {big,desc}[] for stats. IMPORTANT: always provide image_query with English keywords for EVERY content slide so a relevant photo is placed on the slide. Keep content SHORT so it fits without shrinking: stats max 4 items, each big ≤ 8 chars and desc ≤ 28 chars; cards max 3, desc ≤ 60 chars; rows max 5, desc ≤ 70 chars; timeline max 5 steps, title ≤ 18 chars, desc ≤ 40 chars; comparison exactly 2 panels, desc ≤ 120 chars; title ≤ 45 chars. Readability: keep text high-contrast against the background — never use a black background; dark slides use a dark brand color with white text, content slides use a light background with dark text.',
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
        image_side: { type: 'string', enum: ['left', 'right'], description: 'Which side the photo sits on. VARY this across slides for visual rhythm: alternate left/right so the deck does not look monotonous. Default right.' },
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
 * Topic-based theme fallback when the model omits/misspells theme_preset —
 * prevents every deck from defaulting to the same corporate navy.
 */
function deriveThemePreset(title: string): string {
  const t = title.toLowerCase()
  if (/health|medis|kesehatan|dokter|ppds|rs\b|rumah sakit|patient|pasien/.test(t)) return 'forest'
  if (/climate|energy|sustainab|green|lingkungan|karbon|esg|pertanian/.test(t)) return 'forest'
  if (/tech|digital|ai\b|data|software|startup|innovation|teknologi/.test(t)) return 'ocean'
  if (/culture|budaya|sejarah|history|lifestyle|seni|art|wisata|travel/.test(t)) return 'sunset'
  if (/project|pmbok|manajemen|management|pmi|operations|audit|keuangan|finance/.test(t)) return 'slate'
  return 'corporate'
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
      // Curated preset first (recommended path); custom colors only as an explicit override.
      // Unknown/missing preset name → derive from the deck topic instead of always corporate.
      const presetName = String(input.theme_preset ?? '').trim().toLowerCase()
      const preset = THEME_PRESETS[presetName] ?? THEME_PRESETS[deriveThemePreset(title)] ?? DEFAULT_THEME
      currentTheme = { ...preset }
      if (input.theme && typeof input.theme === 'object') {
        const th = input.theme as Record<string, string>
        if (th.primary) currentTheme.primary = th.primary
        if (th.secondary) currentTheme.secondary = th.secondary
        if (th.accent) currentTheme.accent = th.accent
        if (th.background) currentTheme.background = th.background
      }
      const coverQuery = String(input.image_query ?? '').trim() || deriveImageQuery(title, 'title')
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
      const imageSide = input.image_side === 'left' ? 'left' : 'right'

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
        imageSide,
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
