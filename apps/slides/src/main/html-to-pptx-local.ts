import PptxGenJS from 'pptxgenjs'
import type { SlideContent, SlideTheme } from '../renderer/ai/slide-templates'

export const PROVA_JSON_PREFIX = 'prova-json:'

export function encodeSlideMarker(slide: SlideContent): string {
  return PROVA_JSON_PREFIX + Buffer.from(JSON.stringify(slide), 'utf-8').toString('base64')
}

export function decodeSlideMarker(marker: string): SlideContent | null {
  if (!marker.startsWith(PROVA_JSON_PREFIX)) return null
  try {
    const json = Buffer.from(marker.slice(PROVA_JSON_PREFIX.length), 'base64').toString('utf-8')
    return JSON.parse(json) as SlideContent
  } catch {
    return null
  }
}

const DEFAULT_THEME: SlideTheme = {
  primary: '0D2137',
  secondary: '1A73E8',
  accent: '00BFA5',
  background: 'FFFFFF',
  title: '0D2137',
  text: '333333',
}

function hex(c?: string): string {
  return (c ?? '000000').replace(/^#/, '')
}

/** Pick a readable text color (dark or white) for a given solid background hex. */
function textOn(hexColor: string): string {
  const r = parseInt(hexColor.slice(0, 2), 16)
  const g = parseInt(hexColor.slice(2, 4), 16)
  const b = parseInt(hexColor.slice(4, 6), 16)
  const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255
  return lum > 0.6 ? '333333' : 'FFFFFF'
}

/**
 * Text color for elements sitting on a TINT of `base` (fill = base @ ~92% transparency
 * over a light slide). The tint is always light, so the text must be dark — regardless
 * of whether the theme's primary itself is light. Falls back to the theme text color
 * only when that is already dark enough to contrast with the tint.
 */
function textOnTint(base: string, themeText: string): string {
  const dark = '2B3440'
  // Use the theme's text color when it is dark (normal case); otherwise force dark.
  const r = parseInt(themeText.slice(0, 2), 16)
  const g = parseInt(themeText.slice(2, 4), 16)
  const b = parseInt(themeText.slice(4, 6), 16)
  const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255
  return lum < 0.45 ? themeText : dark
}

/** Decode common HTML entities the model sometimes emits (e.g. "&amp;" in titles). */
const HTML_ENTITIES: Array<[string, string]> = [
  ['&lt;', '<'],
  ['&gt;', '>'],
  ['&quot;', '"'],
  ['&#39;', "'"],
  ['&apos;', "'"],
  ['&nbsp;', ' '],
]
function decodeEntities(s: string): string {
  let out = s
  for (const [k, v] of HTML_ENTITIES) out = out.split(k).join(v)
  // &amp; last so "&amp;lt;" decodes once to "&lt;" and not further
  return out.split('&amp;').join('&')
}

/**
 * Largest font size (pt) at which `text` fits a w×h inch box — computed at
 * GENERATION time because PowerPoint does not apply <a:normAutofit> when a file
 * is opened (only when the box is edited), so fit:'shrink' alone renders as
 * overflowing wrapped text in real PowerPoint. Also guarantees the longest
 * single word fits one line (prevents "Biay a mod al" mid-word breaks).
 */
function fitPt(text: string, wIn: number, hIn: number, maxPt: number, minPt = 9): number {
  const CHAR_W = 0.56 // avg glyph width as a fraction of point size (Segoe UI / Carlito)
  const LINE_H = 1.25
  const paragraphs = String(text ?? '').split('\n')
  const maxWord = Math.max(
    1,
    ...paragraphs.flatMap((p) => p.split(/\s+/).map((w) => w.length)),
  )
  const wordPt = Math.floor((wIn * 72) / (CHAR_W * Math.max(maxWord, 1)))
  const start = Math.min(maxPt, wordPt)
  for (let pt = start; pt > minPt; pt--) {
    const cpl = Math.max(1, Math.floor((wIn * 72) / (CHAR_W * pt)))
    let lines = 0
    for (const p of paragraphs) {
      const t = p.trim()
      lines += t ? Math.max(1, Math.ceil(t.length / cpl)) : 1
    }
    if ((lines * pt * LINE_H) / 72 <= hIn) return pt
  }
  return minPt
}

/**
 * Parse a line of text into PptxGenJS runs, honoring inline **bold** markers the
 * model writes. Unmatched markers are stripped so they never render literally.
 */
function parseRuns(line: string, baseColor: string): Array<{ text: string; options: { bold?: boolean; color: string } }> {
  const runs: Array<{ text: string; options: { bold?: boolean; color: string } }> = []
  const re = /\*\*([^*\n]+)\*\*/g
  let last = 0
  for (const m of line.matchAll(re)) {
    const i = m.index ?? 0
    if (i > last) runs.push({ text: line.slice(last, i), options: { color: baseColor } })
    runs.push({ text: m[1] ?? '', options: { bold: true, color: baseColor } })
    last = i + m[0].length
  }
  if (last < line.length) runs.push({ text: line.slice(last), options: { color: baseColor } })
  if (runs.length === 0) runs.push({ text: line, options: { color: baseColor } })
  return runs
}

/** Download an image to a data: URL. Uses Node fetch (follows redirects, has a timeout)
 *  instead of PptxGenJS's https.get, which silently fails on some CDN redirects. */
async function downloadImageToDataUrl(url: string): Promise<string | undefined> {
  try {
    const resp = await fetch(url, { signal: AbortSignal.timeout(15000) })
    if (!resp.ok) return undefined
    const buf = Buffer.from(await resp.arrayBuffer())
    const mime = resp.headers.get('content-type')?.split(';')[0]?.trim() || 'image/jpeg'
    return `data:${mime};base64,${buf.toString('base64')}`
  } catch {
    return undefined
  }
}

import { getSlideBackground } from './slide-backgrounds'

export interface SlideToPptxResult {
  bytes: Uint8Array
  /** Images that could not be downloaded/embedded (page index + url). */
  imageFailures: { page: number; url: string }[]
}

export async function slideContentToPptxBytes(slides: SlideContent[], startNumber = 1): Promise<SlideToPptxResult> {
  const pptx = new PptxGenJS()
  pptx.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 })
  pptx.layout = 'WIDE'
  const imageFailures: { page: number; url: string }[] = []
  // One deck, one font: Segoe UI everywhere (matches the HTML preview template).
  const FONT = { fontFace: 'Segoe UI' }

  for (let si = 0; si < slides.length; si++) {
    const raw = slides[si]!
    // The model sometimes emits HTML entities (&amp; etc.) — decode them so they
    // don't render literally in PowerPoint.
    const slide: SlideContent = {
      ...raw,
      title: decodeEntities(raw.title),
      subtitle: raw.subtitle ? decodeEntities(raw.subtitle) : raw.subtitle,
      eyebrow: raw.eyebrow ? decodeEntities(raw.eyebrow) : raw.eyebrow,
      intro: raw.intro ? decodeEntities(raw.intro) : raw.intro,
      content: (raw.content ?? []).map((c) => {
        if (typeof c === 'string') return decodeEntities(c)
        if (c && typeof c === 'object') {
          const o = { ...(c as Record<string, unknown>) }
          for (const k of ['title', 'desc', 'big']) if (typeof o[k] === 'string') o[k] = decodeEntities(o[k] as string)
          return o
        }
        return c
      }) as SlideContent['content'],
    }
    const slideNumber = startNumber + si
    const th = { ...DEFAULT_THEME, ...(slide.theme ?? {}) }
    const primary = hex(th.primary)
    const secondary = hex(th.secondary)
    const accent = hex(th.accent)
    const bg = hex(th.background)
    const text = hex(th.text)

    const s = pptx.addSlide()
    s.background = { color: bg }

    const isDark = slide.layout === 'section_header' || slide.layout === 'closing' || slide.layout === 'title'
    const titleColor = isDark ? 'FFFFFF' : primary
    const bodyColor = isDark ? 'FFFFFF' : text

    // Hybrid decorative background: gradient/texture rendered from HTML/CSS per
    // theme preset (things PptxGenJS primitives cannot express), embedded as the
    // bottom layer. Text and cards stay native/editable on top of it.
    const bgPng = await getSlideBackground(th, isDark ? 'dark' : 'light')
    if (bgPng) {
      s.addImage({ data: `data:image/png;base64,${bgPng}`, x: 0, y: 0, w: 13.33, h: 7.5 })
    }

    // Resolve the slide photo to a data URL once (used by cover and content slides).
    const imageData = slide.imageUrl ? await downloadImageToDataUrl(slide.imageUrl) : undefined
    if (slide.imageUrl && !imageData) imageFailures.push({ page: si, url: slide.imageUrl })

    if (slide.layout === 'title') {
      if (imageData) {
        s.addImage({ data: imageData, x: 0, y: 0, w: 13.33, h: 7.5, sizing: { type: 'cover', w: 13.33, h: 7.5 } })
        s.addShape('rect', { x: 0, y: 0, w: 13.33, h: 7.5, fill: { color: primary, transparency: 35 } })
      }
      if (slide.eyebrow)
        s.addText(slide.eyebrow.toUpperCase(), { x: 0.7, y: 1.2, w: 11, h: 0.4, fontSize: 16, color: accent, bold: true, charSpacing: 4, ...FONT })
      s.addText(slide.title, { x: 0.7, y: 1.8, w: 11.9, h: 1.4, fontSize: fitPt(slide.title, 11.9, 1.4, 48, 20), color: 'FFFFFF', bold: true, ...FONT })
      if (slide.subtitle)
        s.addText(slide.subtitle, { x: 0.7, y: 3.3, w: 11, h: 0.8, fontSize: fitPt(slide.subtitle, 11, 0.8, 24, 12), color: 'CCCCCC', ...FONT })
      continue
    }

    if (slide.layout === 'closing') {
      s.background = { color: primary }
      s.addText(slide.title || 'Terima Kasih', { x: 0.7, y: 2.8, w: 11.9, h: 1.4, fontSize: fitPt(slide.title || 'Terima Kasih', 11.9, 1.4, 52, 20), color: 'FFFFFF', bold: true, align: 'center', ...FONT })
      if (slide.subtitle)
        s.addText(slide.subtitle, { x: 0.7, y: 4.3, w: 11, h: 0.8, fontSize: fitPt(slide.subtitle, 11, 0.8, 24, 12), color: 'CCCCCC', align: 'center', ...FONT })
      s.addText(String(slideNumber), { x: 12.5, y: 6.95, w: 0.6, h: 0.4, fontSize: 12, color: textOn(primary), align: 'right', ...FONT })
      continue
    }

    if (slide.layout === 'section_header') {
      s.background = { color: primary }
      if (slide.eyebrow)
        s.addText(slide.eyebrow.toUpperCase(), { x: 0.7, y: 1.6, w: 11, h: 0.4, fontSize: 16, color: accent, bold: true, charSpacing: 4, ...FONT })
      s.addText(slide.title, { x: 0.7, y: 2.2, w: 11.9, h: 1.4, fontSize: fitPt(slide.title, 11.9, 1.4, 48, 20), color: 'FFFFFF', bold: true, ...FONT })
      if (slide.intro)
        s.addText(slide.intro, { x: 0.7, y: 3.8, w: 11, h: 0.8, fontSize: fitPt(slide.intro, 11, 0.8, 22, 12), color: 'CCCCCC', ...FONT })
      s.addText(String(slideNumber), { x: 12.5, y: 6.95, w: 0.6, h: 0.4, fontSize: 12, color: textOn(primary), align: 'right', ...FONT })
      continue
    }

    // Content slides
    // Reserve space for the photo only when it ACTUALLY downloaded (imageData) —
    // a failed download must not leave a half-empty slide. Stats deserves the full
    // width: 4 columns beside a photo squeeze descriptions to unreadable sizes.
    const hasImage = !!imageData && slide.layout !== 'stats'
    const imageLeft = slide.imageSide === 'left'
    // Content sits on the opposite side of the photo: left when the image is on
    // the right (legacy), right when the image is on the left.
    const contentX = hasImage && imageLeft ? 5.4 : 0.7
    const contentW = hasImage ? 7.2 : 11.9
    // fit:'shrink' auto-shrinks text on overflow so the initial layout is clean and
    // the QC pass does not need to re-tidy every slide.
    const SHRINK = { fit: 'shrink' as const }
    let y = 0.5
    if (slide.eyebrow) {
      s.addText(slide.eyebrow.toUpperCase(), { x: contentX, y, w: contentW, h: 0.4, fontSize: 14, color: accent, bold: true, charSpacing: 3, ...SHRINK, ...FONT })
      y += 0.45
    }
    s.addText(slide.title, { x: contentX, y, w: contentW, h: 0.8, fontSize: fitPt(slide.title, contentW, 0.8, 34, 16), color: titleColor, bold: true, ...SHRINK, ...FONT })
    y += 0.9
    let introPt = 18
    if (slide.intro) {
      // Taller intro box (1.0") so long intros don't shrink below the bullet size;
      // introPt also clamps the bullets below so the hierarchy never inverts.
      introPt = fitPt(slide.intro, contentW, 1.0, 18, 10)
      s.addText(slide.intro, { x: contentX, y, w: contentW, h: 1.0, fontSize: introPt, color: bodyColor, ...SHRINK, ...FONT })
      y += 1.1
    }
    // Breathing room between the header block and the content below (cards/bullets
    // sat too tight against the intro/title per user feedback).
    y += 0.25

    const content = slide.content
    if (slide.layout === 'cards' && Array.isArray(content) && typeof content[0] === 'object' && 'title' in (content[0] as any)) {
      const cards = content as Array<{ title: string; desc?: string }>
      const cw = hasImage ? 2.2 : 3.9
      const gap = 0.15
      const startX = contentX
      cards.forEach((card, i) => {
        const x = startX + i * (cw + gap)
        // Card without a description shrinks to a compact chip instead of a big empty box.
        const ch = card.desc ? 3.6 : 1.5
        s.addShape('rect', { x, y, w: cw, h: ch, fill: { color: primary, transparency: 92 }, line: { color: primary, transparency: 85, width: 1 }, rectRadius: 0.08 })
        // Card title must contrast with the light card tint — never inherit a light primary.
        s.addText(card.title, { x: x + 0.25, y: y + 0.25, w: cw - 0.5, h: 1.1, fontSize: fitPt(card.title, cw - 0.5, 1.1, hasImage ? 17 : 21, 10), color: textOnTint(primary, text), bold: true, valign: 'top', ...SHRINK, ...FONT })
        if (card.desc)
          s.addText(card.desc, { x: x + 0.25, y: y + 1.45, w: cw - 0.5, h: 2.0, fontSize: fitPt(card.desc, cw - 0.5, 2.0, hasImage ? 12 : 15, 8), color: textOnTint(primary, text), valign: 'top', ...SHRINK, ...FONT })
      })
    } else if (slide.layout === 'stats' && Array.isArray(content) && typeof content[0] === 'object' && 'big' in (content[0] as any)) {
      const stats = content as Array<{ big: string; desc: string }>
      const cw = contentW / stats.length
      // One uniform size across all columns (min of what each needs) — per-item
      // sizing made long labels shrink alone and the row looked uneven. Max 26pt
      // keeps the "big number" emphasis without a drastic jump from the 15pt desc.
      const bigPt = Math.min(...stats.map((st) => fitPt(st.big, cw - 0.2, 1.2, 26, 13)))
      const descPt = Math.min(...stats.map((st) => fitPt(st.desc, cw - 0.2, 1.2, 15, 8)))
      stats.forEach((st, i) => {
        const x = contentX + i * cw
        s.addText(st.big, { x, y, w: cw - 0.2, h: 1.2, fontSize: bigPt, color: secondary, bold: true, align: 'center', ...SHRINK, ...FONT })
        s.addText(st.desc, { x, y: y + 1.3, w: cw - 0.2, h: 1.2, fontSize: descPt, color: text, align: 'center', ...SHRINK, ...FONT })
      })
    } else if (slide.layout === 'rows' && Array.isArray(content) && typeof content[0] === 'object' && 'title' in (content[0] as any)) {
      const rows = content as Array<{ title: string; desc?: string }>
      // Dynamic row pitch: fit all rows between y and the footer (7.0) regardless of count.
      const availH = 7.0 - y
      const rowH = Math.min(0.95, availH / Math.max(rows.length, 1))
      rows.forEach((row, i) => {
        const ry = y + i * rowH
        const hasDesc = !!row.desc
        // Title-only rows use the FULL row height (centered) — the old 50/50 split
        // made them tiny and left a big empty gap under each line.
        const titleH = hasDesc ? rowH * 0.5 : rowH
        const dotY = hasDesc ? ry + 0.12 : ry + rowH / 2 - 0.075
        s.addShape('ellipse', { x: contentX, y: dotY, w: 0.15, h: 0.15, fill: { color: accent } })
        s.addText(row.title, { x: contentX + 0.3, y: ry, w: contentW - 0.3, h: titleH, fontSize: fitPt(row.title, contentW - 0.3, titleH, 18, 12), color: primary, bold: true, valign: hasDesc ? 'top' : 'middle', ...SHRINK, ...FONT })
        if (row.desc)
          s.addText(row.desc, { x: contentX + 0.3, y: ry + rowH * 0.5, w: contentW - 0.3, h: rowH * 0.5, fontSize: fitPt(row.desc, contentW - 0.3, rowH * 0.5, 14, 9), color: text, ...SHRINK, ...FONT })
      })
    } else if (slide.layout === 'agenda' && Array.isArray(content) && typeof content[0] === 'string') {
      ;(content as string[]).forEach((item, i) => {
        // Dynamic spacing so many agenda items never run past the footer.
        const availH = 7.0 - y
        const rowH = Math.min(0.95, availH / Math.max((content as string[]).length, 1))
        const ry = y + i * rowH
        s.addShape('ellipse', { x: contentX, y: ry + 0.05, w: 0.5, h: 0.5, fill: { color: secondary } })
        s.addText(String(i + 1), { x: contentX, y: ry + 0.05, w: 0.5, h: 0.5, fontSize: 16, color: 'FFFFFF', bold: true, align: 'center', valign: 'middle', ...SHRINK, ...FONT })
        s.addText(item, { x: contentX + 0.8, y: ry + 0.05, w: contentW - 0.8, h: 0.5, fontSize: fitPt(item, contentW - 0.8, 0.5, 18, 10), color: text, valign: 'middle', ...SHRINK, ...FONT })
      })
    } else if (slide.layout === 'two_column' && Array.isArray(content) && typeof content[0] === 'string') {
      // Column text may contain literal "• "/"- " bullet lines and **bold** headers
      // (the model writes them). Convert bullets to real bullet paragraphs with
      // hanging indent, and **bold** to bold runs, so styling is consistent.
      // PptxGenJS takes a FLAT run list: each line's runs carry the line's bullet
      // option on its first run, and the line's last run sets breakLine.
      const toRuns = (col: string): Array<{ text: string; options: Record<string, unknown> }> => {
        const out: Array<{ text: string; options: Record<string, unknown> }> = []
        const lines = col.split('\n').map((l) => l.trim()).filter(Boolean)
        lines.forEach((l, li) => {
          const m = /^[•·▪◦‣⁃-]\s+(.*)$/.exec(l)
          const body = m ? (m[1] ?? '') : l
          const runs = parseRuns(body, text)
          runs.forEach((r, ri) => {
            const options: Record<string, unknown> = { ...r.options }
            if (m && ri === 0) options.bullet = { code: '2022', indent: 14 }
            if (ri === runs.length - 1) {
              options.breakLine = true
              if (li < lines.length - 1) options.paraSpaceAfter = 6
            }
            out.push({ text: r.text, options })
          })
        })
        return out
      }
      const [left, right] = content as string[]
      const colW = contentW / 2 - 0.2
      s.addText(toRuns(left), { x: contentX, y, w: colW, h: 4.2, fontSize: fitPt(left, colW, 4.2, 16, 9), color: text, valign: 'top', ...SHRINK, ...FONT })
      s.addText(toRuns(right ?? ''), { x: contentX + contentW / 2 + 0.2, y, w: colW, h: 4.2, fontSize: fitPt(right ?? '', colW, 4.2, 16, 9), color: text, valign: 'top', ...SHRINK, ...FONT })
    } else if (slide.layout === 'timeline' && Array.isArray(content) && typeof content[0] === 'object' && 'title' in (content[0] as any)) {
      // Horizontal timeline: dots on a line, alternating labels above/below.
      const steps = content as Array<{ title: string; desc?: string }>
      const lineY = y + 2.0
      const stepW = contentW / Math.max(steps.length, 1)
      s.addShape('rect', { x: contentX + stepW / 2, y: lineY, w: contentW - stepW, h: 0.03, fill: { color: secondary } })
      steps.forEach((st, i) => {
        const cx = contentX + stepW * i + stepW / 2
        s.addShape('ellipse', { x: cx - 0.09, y: lineY - 0.075, w: 0.18, h: 0.18, fill: { color: accent } })
        const above = i % 2 === 0
        const ty = above ? lineY - 1.55 : lineY + 0.35
        s.addText(st.title, { x: cx - stepW / 2 + 0.1, y: ty, w: stepW - 0.2, h: 0.5, fontSize: fitPt(st.title, stepW - 0.2, 0.5, 15, 9), color: primary, bold: true, align: 'center', ...SHRINK, ...FONT })
        if (st.desc)
          s.addText(st.desc, { x: cx - stepW / 2 + 0.1, y: ty + 0.5, w: stepW - 0.2, h: 1.0, fontSize: fitPt(st.desc, stepW - 0.2, 1.0, 12, 8), color: text, align: 'center', valign: 'top', ...SHRINK, ...FONT })
      })
    } else if (slide.layout === 'quote' && typeof content[0] === 'string') {
      // Full-width pull quote. The decorative quotation glyph sits in its own
      // narrow column left of the quote text — sized to its line so the audit
      // sees no overflow, and horizontally separated so nothing overlaps.
      const quote = content[0] ?? ''
      s.addText('“', { x: contentX, y: y, w: 1.1, h: 1.5, fontSize: 72, color: accent, bold: true, valign: 'top', ...FONT })
      s.addText(quote, { x: contentX + 1.3, y: y, w: contentW - 1.5, h: 2.9, fontSize: fitPt(quote, contentW - 1.5, 2.9, 26, 12), color: primary, italic: true, valign: 'top', ...SHRINK, ...FONT })
      if (slide.intro)
        s.addText(slide.intro, { x: contentX + 1.3, y: y + 3.1, w: contentW - 1.5, h: 0.6, fontSize: fitPt(slide.intro, contentW - 1.5, 0.6, 15, 9), color: text, ...SHRINK, ...FONT })
    } else if (slide.layout === 'big_number' && typeof content[0] === 'string') {
      // One hero number + caption; the intro line acts as the supporting statement.
      s.addText(content[0] ?? '', { x: contentX, y: y + 0.4, w: contentW, h: 2.2, fontSize: fitPt(content[0] ?? '', contentW, 2.2, 110, 40), color: secondary, bold: true, align: 'center', ...SHRINK, ...FONT })
      if (slide.intro)
        s.addText(slide.intro, { x: contentX + 1.2, y: y + 2.9, w: contentW - 2.4, h: 1.0, fontSize: fitPt(slide.intro, contentW - 2.4, 1.0, 20, 10), color: text, align: 'center', valign: 'top', ...SHRINK, ...FONT })
    } else if (slide.layout === 'comparison' && Array.isArray(content) && typeof content[0] === 'object' && 'title' in (content[0] as any)) {
      // Two side-by-side panels (e.g. before/after, problem/solution).
      const panels = content as Array<{ title: string; desc?: string }>
      const pw = (contentW - 0.3) / 2
      panels.slice(0, 2).forEach((p, i) => {
        const x = contentX + i * (pw + 0.3)
        s.addShape('rect', { x, y, w: pw, h: 3.7, fill: { color: primary, transparency: i === 0 ? 94 : 88 }, line: { color: primary, transparency: 80, width: 1 }, rectRadius: 0.08 })
        // Panel text must contrast with the light panel tint — never inherit a light primary.
        s.addText(p.title, { x: x + 0.3, y: y + 0.3, w: pw - 0.6, h: 0.6, fontSize: fitPt(p.title, pw - 0.6, 0.6, 20, 10), color: textOnTint(primary, text), bold: true, ...SHRINK, ...FONT })
        if (p.desc)
          s.addText(p.desc, { x: x + 0.3, y: y + 1.0, w: pw - 0.6, h: 2.5, fontSize: fitPt(p.desc, pw - 0.6, 2.5, 15, 8), color: textOnTint(primary, text), valign: 'top', ...SHRINK, ...FONT })
      })
    } else {
      // Fallback bullets. Items may arrive as objects ({title,desc} / {big,desc})
      // when the model mixes layouts — stringify them as "Title — desc" (title
      // bolded via ** markers) instead of rendering "[object Object]".
      const bullets = (Array.isArray(content) ? content : [])
        .map((c) => {
          if (typeof c === 'string') return c
          if (c && typeof c === 'object') {
            const o = c as Record<string, unknown>
            const t = typeof o.title === 'string' ? o.title.trim() : ''
            const d = typeof o.desc === 'string' ? o.desc.trim() : ''
            const b = typeof o.big === 'string' ? o.big.trim() : ''
            const head = t || b
            if (head && d) return `**${head}** — ${d}`
            return head || d || ''
          }
          return String(c ?? '')
        })
        .filter((s) => s.trim().length > 0)
      // Split embedded newlines into real paragraphs: one bullet per logical line.
      // (A \n inside an item used to render as a bare line break with no bullet,
      // producing the mixed bulleted/unbulleted mess.) Literal "• " prefixes are
      // stripped so they never double up with the real bullet.
      const lines = bullets
        .flatMap((b) => b.split('\n').map((l) => l.trim()).filter(Boolean))
        .map((l) => {
          const m = /^[•·▪◦‣⁃]\s+(.*)$/.exec(l)
          return m ? (m[1] ?? '') : l
        })
      const runs: Array<{ text: string; options: Record<string, unknown> }> = []
      lines.forEach((line, li) => {
        const parsed = parseRuns(line, bodyColor)
        parsed.forEach((r, ri) => {
          const options: Record<string, unknown> = { ...r.options }
          if (ri === 0) options.bullet = { code: '2022', indent: 14 }
          if (ri === parsed.length - 1) {
            options.breakLine = true
            if (li < lines.length - 1) options.paraSpaceAfter = 9
          }
          runs.push({ text: r.text, options })
        })
      })
      // Bullet body 15pt max, clamped to the intro size so it never outgrows it.
      s.addText(runs, { x: contentX, y, w: contentW, h: 3.8, fontSize: Math.min(fitPt(lines.join('\n'), contentW, 3.8, 15, 9), introPt), color: bodyColor, valign: 'top', ...SHRINK, ...FONT })
    }

    // Footer
    s.addShape('rect', { x: 0, y: 7.1, w: 13.33, h: 0.4, fill: { color: primary } })
    s.addText(String(slideNumber), { x: 12.4, y: 7.1, w: 0.7, h: 0.4, fontSize: 11, color: textOn(primary), align: 'right', valign: 'middle', ...FONT })

    // Optional image (Pixabay) — right side by default, or left when imageSide='left'
    if (imageData && hasImage) {
      const imgX = slide.imageSide === 'left' ? 0.7 : 8.2
      s.addImage({ data: imageData, x: imgX, y: 1.2, w: 4.4, h: 5.4, sizing: { type: 'cover', w: 4.4, h: 5.4 } })
    }
  }

  const buf = await pptx.write({ outputType: 'arraybuffer' })
  return { bytes: new Uint8Array(buf as ArrayBuffer), imageFailures }
}
