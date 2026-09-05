/**
 * Hybrid slide backgrounds: each theme preset gets a decorative background
 * rendered from HTML/CSS (gradients, soft blobs, subtle grid — things PptxGenJS
 * primitives cannot express) into a PNG via an offscreen BrowserWindow, then
 * embedded as the slide's bottom layer. Text and cards stay native/editable.
 *
 * One PNG per theme (cached per session); 2 variants: dark (cover/section/closing)
 * and light (content slides).
 */
import { BrowserWindow } from 'electron'
import type { SlideTheme } from '../renderer/ai/slide-templates'

export type BackgroundVariant = 'dark' | 'light'

const W_PX = 1600
const H_PX = 900

interface BgSpec {
  dark: string
  light: string
}

/** Decorative CSS per theme preset, keyed by the preset's primary hue family. */
const SPECS: Record<string, BgSpec> = {
  corporate: {
    dark: `
      background:
        radial-gradient(900px 600px at 85% -10%, rgba(26,115,232,0.35), transparent 60%),
        radial-gradient(700px 500px at -10% 110%, rgba(0,191,165,0.22), transparent 55%),
        linear-gradient(135deg, #0D2137 0%, #123252 55%, #0D2137 100%);`,
    light: `
      background:
        radial-gradient(800px 500px at 110% -5%, rgba(26,115,232,0.10), transparent 60%),
        radial-gradient(600px 400px at -5% 105%, rgba(0,191,165,0.08), transparent 55%),
        linear-gradient(180deg, #FFFFFF 0%, #F4F7FB 100%);`,
  },
  ocean: {
    dark: `
      background:
        radial-gradient(900px 600px at 80% -10%, rgba(50,140,193,0.40), transparent 60%),
        radial-gradient(700px 500px at -10% 110%, rgba(242,161,4,0.18), transparent 55%),
        linear-gradient(135deg, #0B3C5D 0%, #14507A 55%, #0B3C5D 100%);`,
    light: `
      background:
        radial-gradient(800px 500px at 110% -5%, rgba(50,140,193,0.12), transparent 60%),
        radial-gradient(600px 400px at -5% 105%, rgba(242,161,4,0.08), transparent 55%),
        linear-gradient(180deg, #FFFFFF 0%, #F3F8FB 100%);`,
  },
  forest: {
    dark: `
      background:
        radial-gradient(900px 600px at 85% -10%, rgba(45,106,79,0.45), transparent 60%),
        radial-gradient(700px 500px at -10% 110%, rgba(233,196,106,0.16), transparent 55%),
        linear-gradient(135deg, #1B4332 0%, #25573F 55%, #1B4332 100%);`,
    light: `
      background:
        radial-gradient(800px 500px at 110% -5%, rgba(45,106,79,0.10), transparent 60%),
        radial-gradient(600px 400px at -5% 105%, rgba(233,196,106,0.10), transparent 55%),
        linear-gradient(180deg, #FFFFFF 0%, #F5F9F6 100%);`,
  },
  sunset: {
    dark: `
      background:
        radial-gradient(900px 600px at 85% -10%, rgba(199,81,70,0.38), transparent 60%),
        radial-gradient(700px 500px at -10% 110%, rgba(244,162,89,0.20), transparent 55%),
        linear-gradient(135deg, #5D2A42 0%, #7A3A57 55%, #5D2A42 100%);`,
    light: `
      background:
        radial-gradient(800px 500px at 110% -5%, rgba(199,81,70,0.10), transparent 60%),
        radial-gradient(600px 400px at -5% 105%, rgba(244,162,89,0.10), transparent 55%),
        linear-gradient(180deg, #FFFFFF 0%, #FBF6F3 100%);`,
  },
  slate: {
    dark: `
      background:
        radial-gradient(900px 600px at 85% -10%, rgba(82,121,111,0.42), transparent 60%),
        radial-gradient(700px 500px at -10% 110%, rgba(224,122,95,0.18), transparent 55%),
        linear-gradient(135deg, #2F3E46 0%, #3C5049 55%, #2F3E46 100%);`,
    light: `
      background:
        radial-gradient(800px 500px at 110% -5%, rgba(82,121,111,0.10), transparent 60%),
        radial-gradient(600px 400px at -5% 105%, rgba(224,122,95,0.08), transparent 55%),
        linear-gradient(180deg, #FFFFFF 0%, #F5F7F6 100%);`,
  },
}

/** Subtle dot grid overlay shared by all themes (adds texture without noise). */
const GRID_OVERLAY = `
  <div style="position:absolute;inset:0;background-image:radial-gradient(rgba(255,255,255,0.05) 1px, transparent 1px);background-size:28px 28px;"></div>`

function bgHtml(spec: BgSpec, variant: BackgroundVariant): string {
  const css = variant === 'dark' ? spec.dark : spec.light
  return `<!doctype html><html><head><style>
    html,body{margin:0;padding:0;width:${W_PX}px;height:${H_PX}px;overflow:hidden}
    .stage{position:relative;width:${W_PX}px;height:${H_PX}px;${css}}
    .stage > *{position:absolute}
  </style></head><body><div class="stage">${GRID_OVERLAY}</div></body></html>`
}

/** Render one background PNG (base64, no data: prefix) via a hidden window. */
async function renderBackgroundPng(spec: BgSpec, variant: BackgroundVariant): Promise<Buffer> {
  const win = new BrowserWindow({
    show: false,
    width: W_PX,
    height: H_PX,
    webPreferences: { sandbox: true, offscreen: true },
  })
  try {
    const html = bgHtml(spec, variant)
    await win.loadURL('data:text/html;base64,' + Buffer.from(html, 'utf8').toString('base64'))
    await win.webContents.executeJavaScript('document.fonts.ready.then(() => undefined)', true)
    const image = await win.webContents.capturePage({ x: 0, y: 0, width: W_PX, height: H_PX })
    return image.toPNG()
  } finally {
    if (!win.isDestroyed()) win.destroy()
  }
}

const cache = new Map<string, string>()

/** Theme preset name → spec key (falls back to 'corporate' styling for custom themes). */
export function presetKeyForTheme(theme: SlideTheme): string {
  for (const key of Object.keys(SPECS)) {
    const preset = THEME_PRESET_PRIMARY[key]
    if (preset && preset.toLowerCase() === theme.primary.toLowerCase()) return key
  }
  return 'corporate'
}

const THEME_PRESET_PRIMARY: Record<string, string> = {
  corporate: '#0D2137',
  ocean: '#0B3C5D',
  forest: '#1B4332',
  sunset: '#5D2A42',
  slate: '#2F3E46',
}

/** Get (or render+cache) the decorative background PNG as base64 for a theme. */
export async function getSlideBackground(
  theme: SlideTheme,
  variant: BackgroundVariant,
): Promise<string | undefined> {
  const key = presetKeyForTheme(theme)
  const cacheKey = `${key}:${variant}`
  const hit = cache.get(cacheKey)
  if (hit) return hit
  const spec = SPECS[key] ?? SPECS.corporate!
  try {
    const png = await renderBackgroundPng(spec, variant)
    const b64 = png.toString('base64')
    cache.set(cacheKey, b64)
    return b64
  } catch (err) {
    console.warn('[slides] decorative background render failed; falling back to solid fill:', err)
    return undefined
  }
}