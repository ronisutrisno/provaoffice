/**
 * AI slide generation — shared types + house-style color sets.
 *
 * Templates were ported from the proxsis-llm-fastapi PPTX generator
 * (house style: Georgia titles, rounded boxes, banner KEY MESSAGE, uniform
 * cards, business_case variants). Rendering happens in
 * `src/main/html-to-pptx-local.ts` (PptxGenJS, native editable elements).
 * This module only defines the data contract + the three curated color sets.
 */

export interface SlideTheme {
  /** brand primary (maroon / navy / green) — titles, footer, dark text on light */
  primary: string
  /** brand light shade — accents on dark panels */
  secondary: string
  /** small-caps eyebrow / dots / markers */
  accent: string
  /** slide background (warm off-white) */
  background: string
  title: string
  text: string
  footerText?: string
  // ── house-style additions (optional; renderer derives fallbacks) ──
  /** darkest brand shade: dark cards, cover/divider/closing backgrounds */
  primaryDark?: string
  /** mid brand shade: banner accent bar, labels on dark cards */
  primaryLight?: string
  /** KEY MESSAGE banner + numbered-list zebra fill */
  bannerBg?: string
  zebra?: string
  /** near-black body ink */
  ink?: string
  /** warm gray body/muted text */
  muted?: string
}

/** Resolve a house-style theme with derived fallbacks for older payloads. */
export function houseTheme(th: SlideTheme): Required<Omit<SlideTheme, 'footerText'>> & { footerText: string } {
  return {
    primary: th.primary,
    secondary: th.secondary,
    accent: th.accent,
    background: th.background,
    title: th.title,
    text: th.text,
    footerText: th.footerText ?? th.muted ?? th.text,
    primaryDark: th.primaryDark ?? th.primary,
    primaryLight: th.primaryLight ?? th.secondary,
    bannerBg: th.bannerBg ?? th.background,
    zebra: th.zebra ?? th.bannerBg ?? th.background,
    ink: th.ink ?? th.title,
    muted: th.muted ?? th.text,
  }
}

export const DEFAULT_THEME: SlideTheme = {
  primary: '#5C0000',
  secondary: '#822828',
  accent: '#822828',
  background: '#FAF8F7',
  title: '#1F1A1A',
  text: '#6E5A5A',
  footerText: '#6E5A5A',
  primaryDark: '#3A0000',
  primaryLight: '#822828',
  bannerBg: '#F3EFEF',
  zebra: '#F5F0F0',
  ink: '#1F1A1A',
  muted: '#6E5A5A',
}

/**
 * Three curated color sets — same house style, different brand hue.
 * All contrast-checked: white text on primary/primaryDark, dark ink on light bg.
 */
export const THEME_PRESETS: Record<string, SlideTheme> = {
  maroon: { ...DEFAULT_THEME },
  navy: {
    primary: '#0F2B3C',
    secondary: '#3E7A96',
    accent: '#2E5A70',
    background: '#F7FAFC',
    title: '#152530',
    text: '#5C6B75',
    footerText: '#5C6B75',
    primaryDark: '#081C28',
    primaryLight: '#2E5A70',
    bannerBg: '#EAF0F3',
    zebra: '#F0F5F7',
    ink: '#152530',
    muted: '#5C6B75',
  },
  green: {
    primary: '#1B4332',
    secondary: '#40916C',
    accent: '#2D6A4F',
    background: '#F7FAF8',
    title: '#1C2620',
    text: '#5C6E62',
    footerText: '#5C6E62',
    primaryDark: '#12301F',
    primaryLight: '#2D6A4F',
    bannerBg: '#EAF2EC',
    zebra: '#F0F6F1',
    ink: '#1C2620',
    muted: '#5C6E62',
  },
}

export const THEME_PRESET_NAMES = Object.keys(THEME_PRESETS)

// ── Content contract ─────────────────────────────────────────────────────

/** business_case payload — mirrors the fastapi generator's dict schema. */
export interface BusinessCasePayload {
  /** 'full' (default) | 'solution' | 'metrics' */
  variant?: 'full' | 'solution' | 'metrics'
  /** @deprecated badge pill kanan-atas dihapus 2026-09-13; field diterima tapi tidak dirender */
  badge?: string
  key_message?: string
  background?: string
  business_problem?: string
  solution?: string[] | string
  table_title?: string
  table?: { headers: string[]; rows: string[][] }
  /** variant=solution: up to 3 benefit cards */
  benefits?: Array<{ title: string; desc?: string }>
  /** variant=metrics: up to 3 big numbers */
  metrics?: Array<{ big: string; desc?: string }>
  challenge?: string
  impact?: string
  // ── Judul kotak dinamis (opsional; default label house-style) ──
  /** default "BACKGROUND" */
  background_label?: string
  /** default "BUSINESS PROBLEM" */
  problem_label?: string
  /** default "PROPOSED SOLUTION" */
  solution_label?: string
  /** default "TANTANGAN" */
  challenge_label?: string
  /** default "DAMPAK" */
  impact_label?: string
}

export interface SlideContent {
  title: string
  subtitle?: string
  eyebrow?: string
  intro?: string
  content?: string[] | Array<{ title: string; desc?: string }> | Array<{ big: string; desc: string }>
  /**
   * House-style layouts: numbered_list, cards, stats, callout, agenda,
   * business_case, section_header, closing, blank.
   * Legacy keys (title_content/two_column/rows/timeline/quote/big_number/
   * comparison) are still accepted and redirected by the renderer/tools.
   */
  layout:
    | 'title'
    | 'numbered_list'
    | 'cards'
    | 'stats'
    | 'callout'
    | 'agenda'
    | 'business_case'
    | 'section_header'
    | 'closing'
    | 'blank'
    // legacy (redirected)
    | 'title_content'
    | 'two_column'
    | 'rows'
    | 'timeline'
    | 'quote'
    | 'big_number'
    | 'comparison'
    | 'section_header_legacy'
  imageUrl?: string
  /** which side the photo sits on; undefined = right (legacy default) */
  imageSide?: 'left' | 'right'
  theme?: Partial<SlideTheme>
  /** business_case payload (layout='business_case') */
  bc?: BusinessCasePayload
}

/** Map legacy layout names to their house-style successors (silent redirect). */
export function redirectLayout(layout: string): SlideContent['layout'] {
  switch (layout) {
    case 'title_content':
    case 'two_column':
    case 'rows':
    case 'timeline':
    case 'comparison':
      return 'numbered_list'
    case 'quote':
      return 'callout'
    case 'big_number':
      return 'stats'
    case 'section_header_legacy':
      return 'section_header'
    default:
      return layout as SlideContent['layout']
  }
}
