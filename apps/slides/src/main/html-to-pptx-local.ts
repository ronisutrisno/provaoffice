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

export async function slideContentToPptxBytes(slides: SlideContent[]): Promise<Uint8Array> {
  const pptx = new PptxGenJS()
  pptx.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 })
  pptx.layout = 'WIDE'

  for (const slide of slides) {
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

    if (slide.layout === 'title') {
      s.background = { color: primary }
      if (slide.imageUrl) {
        try {
          s.addImage({ path: slide.imageUrl, x: 0, y: 0, w: 13.33, h: 7.5, sizing: { type: 'cover', w: 13.33, h: 7.5 } })
          s.addShape('rect', { x: 0, y: 0, w: 13.33, h: 7.5, fill: { color: primary, transparency: 35 } })
        } catch {
          /* image failed — solid background stays */
        }
      }
      if (slide.eyebrow)
        s.addText(slide.eyebrow.toUpperCase(), { x: 0.7, y: 1.2, w: 11, h: 0.4, fontSize: 16, color: accent, bold: true, charSpacing: 4 })
      s.addText(slide.title, { x: 0.7, y: 1.8, w: 11.9, h: 1.4, fontSize: 48, color: 'FFFFFF', bold: true })
      if (slide.subtitle)
        s.addText(slide.subtitle, { x: 0.7, y: 3.3, w: 11, h: 0.8, fontSize: 24, color: 'CCCCCC' })
      continue
    }

    if (slide.layout === 'closing') {
      s.background = { color: primary }
      s.addText(slide.title || 'Terima Kasih', { x: 0.7, y: 2.8, w: 11.9, h: 1.4, fontSize: 52, color: 'FFFFFF', bold: true, align: 'center' })
      if (slide.subtitle)
        s.addText(slide.subtitle, { x: 0.7, y: 4.3, w: 11, h: 0.8, fontSize: 24, color: 'CCCCCC', align: 'center' })
      continue
    }

    if (slide.layout === 'section_header') {
      s.background = { color: primary }
      if (slide.eyebrow)
        s.addText(slide.eyebrow.toUpperCase(), { x: 0.7, y: 1.6, w: 11, h: 0.4, fontSize: 16, color: accent, bold: true, charSpacing: 4 })
      s.addText(slide.title, { x: 0.7, y: 2.2, w: 11.9, h: 1.4, fontSize: 48, color: 'FFFFFF', bold: true })
      if (slide.intro)
        s.addText(slide.intro, { x: 0.7, y: 3.8, w: 11, h: 0.8, fontSize: 22, color: 'CCCCCC' })
      continue
    }

    // Content slides
    const hasImage = !!slide.imageUrl
    const contentW = hasImage ? 7.2 : 11.9
    let y = 0.5
    if (slide.eyebrow) {
      s.addText(slide.eyebrow.toUpperCase(), { x: 0.7, y, w: contentW, h: 0.4, fontSize: 14, color: accent, bold: true, charSpacing: 3 })
      y += 0.45
    }
    s.addText(slide.title, { x: 0.7, y, w: contentW, h: 0.8, fontSize: 32, color: titleColor, bold: true })
    y += 0.85
    if (slide.intro) {
      s.addText(slide.intro, { x: 0.7, y, w: contentW, h: 0.7, fontSize: 16, color: bodyColor })
      y += 0.75
    }

    const content = slide.content
    if (slide.layout === 'cards' && Array.isArray(content) && typeof content[0] === 'object' && 'title' in (content[0] as any)) {
      const cards = content as Array<{ title: string; desc?: string }>
      const cw = hasImage ? 2.2 : 3.9
      const gap = 0.15
      const startX = 0.7
      cards.forEach((card, i) => {
        const x = startX + i * (cw + gap)
        s.addShape('rect', { x, y, w: cw, h: 3.6, fill: { color: primary, transparency: 92 }, line: { color: primary, transparency: 85, width: 1 }, rectRadius: 0.08 })
        s.addShape('ellipse', { x: x + 0.25, y: y + 0.25, w: 0.16, h: 0.16, fill: { color: accent } })
        s.addText(card.title, { x: x + 0.25, y: y + 0.55, w: cw - 0.5, h: 0.9, fontSize: hasImage ? 16 : 20, color: primary, bold: true, valign: 'top' })
        if (card.desc)
          s.addText(card.desc, { x: x + 0.25, y: y + 1.5, w: cw - 0.5, h: 2.0, fontSize: hasImage ? 12 : 15, color: text, valign: 'top' })
      })
    } else if (slide.layout === 'stats' && Array.isArray(content) && typeof content[0] === 'object' && 'big' in (content[0] as any)) {
      const stats = content as Array<{ big: string; desc: string }>
      const cw = contentW / stats.length
      stats.forEach((st, i) => {
        const x = 0.7 + i * cw
        s.addText(st.big, { x, y, w: cw - 0.2, h: 1.2, fontSize: 44, color: secondary, bold: true, align: 'center' })
        s.addText(st.desc, { x, y: y + 1.3, w: cw - 0.2, h: 1.2, fontSize: 15, color: text, align: 'center' })
      })
    } else if (slide.layout === 'rows' && Array.isArray(content) && typeof content[0] === 'object' && 'title' in (content[0] as any)) {
      const rows = content as Array<{ title: string; desc?: string }>
      rows.forEach((row, i) => {
        const ry = y + i * 0.95
        s.addShape('ellipse', { x: 0.7, y: ry + 0.15, w: 0.15, h: 0.15, fill: { color: accent } })
        s.addText(row.title, { x: 1.0, y: ry, w: contentW - 0.3, h: 0.45, fontSize: 18, color: primary, bold: true })
        if (row.desc)
          s.addText(row.desc, { x: 1.0, y: ry + 0.45, w: contentW - 0.3, h: 0.45, fontSize: 14, color: text })
      })
    } else if (slide.layout === 'agenda' && Array.isArray(content) && typeof content[0] === 'string') {
      ;(content as string[]).forEach((item, i) => {
        const ry = y + i * 0.95
        s.addShape('ellipse', { x: 0.7, y: ry + 0.05, w: 0.5, h: 0.5, fill: { color: secondary } })
        s.addText(String(i + 1), { x: 0.7, y: ry + 0.05, w: 0.5, h: 0.5, fontSize: 16, color: 'FFFFFF', bold: true, align: 'center', valign: 'middle' })
        s.addText(item, { x: 1.5, y: ry + 0.05, w: contentW - 0.8, h: 0.5, fontSize: 18, color: text, valign: 'middle' })
      })
    } else if (slide.layout === 'two_column' && Array.isArray(content) && typeof content[0] === 'string') {
      const [left, right] = content as string[]
      s.addText(left, { x: 0.7, y, w: contentW / 2 - 0.2, h: 4.2, fontSize: 16, color: text, valign: 'top' })
      s.addText(right ?? '', { x: 0.7 + contentW / 2 + 0.2, y, w: contentW / 2 - 0.2, h: 4.2, fontSize: 16, color: text, valign: 'top' })
    } else {
      const bullets = Array.isArray(content) ? content.map((c) => String(c)) : []
      s.addText(bullets.map((b) => ({ text: b, options: { bullet: { code: '2022' }, paraSpaceAfter: 9 } })), { x: 0.7, y, w: contentW, h: 4.2, fontSize: 18, color: bodyColor, valign: 'top' })
    }

    // Footer
    s.addShape('rect', { x: 0, y: 7.1, w: 13.33, h: 0.4, fill: { color: primary } })

    // Optional image (Pixabay) — right side, 4.5 wide area (content narrows when image present)
    if (slide.imageUrl) {
      try {
        s.addImage({ path: slide.imageUrl, x: 8.2, y: 1.2, w: 4.4, h: 5.4, sizing: { type: 'cover', w: 4.4, h: 5.4 } })
      } catch {
        /* image download failed — skip silently */
      }
    }
  }

  const buf = await pptx.write({ outputType: 'arraybuffer' })
  return new Uint8Array(buf as ArrayBuffer)
}
