import PptxGenJS from 'pptxgenjs'
import {
  DEFAULT_THEME,
  houseTheme,
  redirectLayout,
  type BusinessCasePayload,
  type SlideContent,
} from '../renderer/ai/slide-templates'

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
 * single word fits one line.
 */
function fitPt(text: string, wIn: number, hIn: number, maxPt: number, minPt = 9): number {
  const CHAR_W = 0.56 // avg glyph width as a fraction of point size (Georgia)
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
 * Font floor + truncation: instead of shrinking long text to micro size, cap it
 * at `minPt` and cut to the number of lines that actually fit (ellipsis).
 * Mirrors the fastapi generator's desc-truncation behavior.
 */
function clampFit(
  text: string,
  wIn: number,
  hIn: number,
  maxPt = 12,
  minPt = 11,
): { text: string; pt: number } {
  const pt = fitPt(text, wIn, hIn, maxPt, minPt)
  if (pt > minPt) return { text, pt }
  const cpl = Math.max(10, Math.floor((wIn * 72) / (0.56 * minPt)))
  const maxLines = Math.max(1, Math.floor((hIn * 72) / (minPt * 1.25)))
  const maxChars = cpl * maxLines
  let t = text
  if (t.length > maxChars) t = `${t.slice(0, Math.max(0, maxChars - 1)).trimEnd()}…`
  return { text: t, pt: minPt }
}

/** Download an image to a data: URL (Node fetch follows CDN redirects; PptxGenJS's does not). */
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

export interface SlideToPptxResult {
  bytes: Uint8Array
  /** Images that could not be downloaded/embedded (page index + url). */
  imageFailures: { page: number; url: string }[]
}

// ── House-style geometry (ported from proxsis-llm-fastapi pptx_tools.py) ──
// Slide 13.33 x 7.5 in; margin 0.6; content width 12.13.
const W = 13.33
const MX = 0.6
const CW = W - 2 * MX
const FOOTER_Y = 7.02

const SHRINK = { fit: 'shrink' as const }

export async function slideContentToPptxBytes(
  slides: SlideContent[],
  startNumber = 1,
): Promise<SlideToPptxResult> {
  const pptx = new PptxGenJS()
  pptx.defineLayout({ name: 'WIDE', width: W, height: 7.5 })
  pptx.layout = 'WIDE'
  const imageFailures: { page: number; url: string }[] = []
  // One deck, one font: Georgia everywhere (titles AND body) — house style.
  const FONT = { fontFace: 'Georgia' }

  for (let si = 0; si < slides.length; si++) {
    const raw = slides[si]!
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
          for (const k of ['title', 'desc', 'big'])
            if (typeof o[k] === 'string') o[k] = decodeEntities(o[k] as string)
          return o
        }
        return c
      }) as SlideContent['content'],
    }
    const layout = redirectLayout(slide.layout)
    const th = houseTheme({ ...DEFAULT_THEME, ...(slide.theme ?? {}) })
    const P = hex(th.primary)
    const PL = hex(th.primaryLight)
    const PD = hex(th.primaryDark)
    const BG = hex(th.background)
    const INK = hex(th.ink)
    const MUT = hex(th.muted)
    const ACC = hex(th.accent)
    const BANNER = hex(th.bannerBg)
    const ZEBRA = hex(th.zebra)
    const WHITE = 'FFFFFF'
    // Bar/dot color cycle (mirrors _COMPONENT_PALETTE)
    const PAL = [P, PL, ACC, hex(th.secondary)]
    const pal = (i: number) => PAL[i % PAL.length]

    const s = pptx.addSlide()
    s.background = { color: BG }
    const slideNumber = startNumber + si

    const imageData = slide.imageUrl ? await downloadImageToDataUrl(slide.imageUrl) : undefined
    if (slide.imageUrl && !imageData) imageFailures.push({ page: si, url: slide.imageUrl })

    const addText = (
      text: string,
      opts: Parameters<typeof s.addText>[1],
    ) => s.addText(text, { ...FONT, ...SHRINK, ...opts })

    // ══════════ COVER (title) — maroon + photo panel kanan 38% ══════════
    if (layout === 'title') {
      s.background = { color: P }
      const hasPhoto = !!imageData
      const textW = hasPhoto ? W * 0.62 - 1.7 : W - 1.8
      if (hasPhoto) {
        const pw = W * 0.38
        s.addImage({
          data: imageData!,
          x: W - pw,
          y: 0,
          w: pw,
          h: 7.5,
          sizing: { type: 'cover', w: pw, h: 7.5 },
        })
        s.addShape('rect', { x: W - pw - 0.03, y: 0, w: 0.06, h: 7.5, fill: { color: PL } })
      }
      addText((slide.eyebrow || 'PRESENTASI').toUpperCase(), {
        x: 0.9, y: 1.15, w: 6, h: 0.4, fontSize: 13, color: 'D9D9D9', bold: true, charSpacing: 3,
        fit: 'none',
      })
      s.addShape('rect', { x: 0.9, y: 1.72, w: 1.5, h: 0.06, fill: { color: 'D9D9D9' } })
      addText(slide.title, {
        x: 0.9, y: 2.45, w: textW, h: 2.5, fontSize: fitPt(slide.title, textW, 2.5, 44, 18),
        color: WHITE, bold: true, valign: 'middle', lineSpacingMultiple: 1.1,
      })
      if (slide.subtitle)
        addText(slide.subtitle, {
          x: 0.9, y: 5.2, w: textW, h: 0.95,
          fontSize: fitPt(slide.subtitle, textW, 0.95, 17, 10), color: 'D9CFCF',
          lineSpacingMultiple: 1.25,
        })
      continue
    }

    // ══════════ SECTION DIVIDER — maroon + photo panel sisi (imageSide) ══════════
    if (layout === 'section_header') {
      s.background = { color: P }
      const hasPhoto = !!imageData
      const left = slide.imageSide !== 'left' ? false : true // default photo kanan
      const pw = hasPhoto ? W * 0.42 : 0
      if (hasPhoto) {
        const px = left ? 0 : W - pw
        s.addImage({
          data: imageData!,
          x: px,
          y: 0,
          w: pw,
          h: 7.5,
          sizing: { type: 'cover', w: pw, h: 7.5 },
        })
      }
      const tx = hasPhoto && !left ? 0.9 : hasPhoto && left ? pw + 0.6 : MX
      const tw = hasPhoto ? W - pw - tx - 0.6 : CW
      addText('SECTION', {
        x: tx, y: 1.2, w: Math.min(tw, 4), h: 0.4, fontSize: 13, color: 'D9D9D9', bold: true,
        charSpacing: 3, fit: 'none',
      })
      s.addShape('rect', { x: tx, y: 1.75, w: 1.5, h: 0.06, fill: { color: 'D9D9D9' } })
      if (slide.eyebrow)
        addText(slide.eyebrow, {
          x: tx, y: 2.35, w: tw, h: 1.1, fontSize: fitPt(slide.eyebrow, tw, 1.1, 60, 24),
          color: WHITE, bold: true,
        })
      addText(slide.title, {
        x: tx, y: 3.7, w: tw, h: 2.2, fontSize: fitPt(slide.title, tw, 2.2, 40, 18),
        color: WHITE, bold: true, lineSpacingMultiple: 1.1,
      })
      continue
    }

    // ══════════ CLOSING ══════════
    if (layout === 'closing') {
      s.background = { color: PD }
      s.addShape('rect', { x: W / 2 - 1.0, y: 2.2, w: 2.0, h: 0.06, fill: { color: 'D9D9D9' } })
      addText(slide.title || 'Terima Kasih', {
        x: MX, y: 2.55, w: CW, h: 2.35, fontSize: fitPt(slide.title || 'Terima Kasih', CW, 2.35, 40, 18),
        color: WHITE, bold: true, align: 'center', valign: 'middle', lineSpacingMultiple: 1.1,
      })
      if (slide.subtitle)
        addText(slide.subtitle, {
          x: MX + 1.5, y: 5.15, w: CW - 3, h: 0.9,
          fontSize: fitPt(slide.subtitle, CW - 3, 0.9, 17, 10), color: 'D9CFCF', align: 'center',
          lineSpacingMultiple: 1.25,
        })
      continue
    }

    // ══════════ BLANK ══════════
    if (layout === 'blank') continue

    // ══════════ BUSINESS CASE (3 varian) ══════════
    if (layout === 'business_case') {
      const bc: BusinessCasePayload = slide.bc ?? {}
      const variant = (bc.variant ?? 'full').toLowerCase()
      // Judul kotak dinamis (default label house-style) — 2026-09-13
      const lbl = (v: string | undefined, dflt: string): string => String(v ?? '').trim() || dflt
      const L_BG = lbl(bc.background_label, 'BACKGROUND')
      const L_PROB = lbl(bc.problem_label, 'BUSINESS PROBLEM')
      const L_SOL = lbl(bc.solution_label, 'PROPOSED SOLUTION')
      const L_CHAL = lbl(bc.challenge_label, 'TANTANGAN')
      const L_IMP = lbl(bc.impact_label, 'DAMPAK')
      // Eyebrow + judul di slot SAMA seperti slide konten lain (konsisten antar layout).
      // Badge pill kanan-atas DIHAPUS 2026-09-13 (judul panjang tertutup badge).
      let hy = 0.55
      if (slide.eyebrow) {
        addText(slide.eyebrow.toUpperCase(), {
          x: MX, y: hy, w: CW, h: 0.35, fontSize: 12, color: ACC, bold: true, charSpacing: 2,
          fit: 'none',
        })
        hy += 0.42
      }
      addText(slide.title, {
        x: MX, y: hy, w: CW, h: 0.85,
        fontSize: fitPt(slide.title, CW, 0.85, 26, 14), color: PD, bold: true,
        valign: 'middle',
      })
      // Banner KEY MESSAGE — HANYA bila isinya ada (dulu frame selalu digambar
      // sehingga payload tanpa key_message menghasilkan pita kosong).
      const km = String(bc.key_message ?? '').trim()
      const by = 1.78
      if (km) {
        s.addShape('roundRect', { x: MX, y: by, w: CW, h: 0.78, fill: { color: BANNER }, rectRadius: 0.04, line: { type: 'none' } })
        s.addShape('rect', { x: MX, y: by, w: 0.09, h: 0.78, fill: { color: PL } })
        addText('KEY MESSAGE', { x: MX + 0.3, y: by + 0.24, w: 1.75, h: 0.3, fontSize: 11, color: P, bold: true, fit: 'none' })
        addText(km, {
          x: MX + 2.1, y: by + 0.03, w: CW - 2.4, h: 0.72,
          fontSize: fitPt(km, CW - 2.4, 0.72, 13, 9), color: PD, bold: true,
          valign: 'middle', lineSpacingMultiple: 1.25,
        })
      }
      // Batas atas konten varian: slot banner dipakai kembali bila banner tidak ada.
      const yTop = km ? 2.68 : 1.85

      if (variant === 'solution') {
        // Kartu SOLUSI full-width gelap + 3 benefit putih
        const sy = yTop
        const shh = 1.95
        s.addShape('roundRect', { x: MX, y: sy, w: CW, h: shh, fill: { color: PD }, rectRadius: 0.04, line: { type: 'none' } })
        addText(L_SOL, { x: MX + 0.3, y: sy + 0.16, w: CW - 0.6, h: 0.28, fontSize: 11, color: WHITE, bold: true, fit: 'none' })
        const paras = (Array.isArray(bc.solution) ? bc.solution : [String(bc.solution ?? '')])
          .map((p) => String(p).trim()).filter(Boolean).slice(0, 3)
        const segH = paras.length ? (shh - 0.55) / paras.length : 0
        paras.forEach((p, i) => {
          const fit = clampFit(p, CW - 0.6, segH - 0.04)
          addText(fit.text, {
            x: MX + 0.3, y: sy + 0.5 + i * segH, w: CW - 0.6, h: segH - 0.04,
            fontSize: fit.pt, color: 'F0EAEA', lineSpacingMultiple: 1.3,
          })
        })
        const benefits = (bc.benefits ?? []).slice(0, 3)
        if (benefits.length) {
          const by2 = sy + shh + 0.2
          const bh = Math.min(2.15, Math.max(1.2, FOOTER_Y - by2 - 0.12))
          const gap = 0.23
          const bw = (CW - gap * (benefits.length - 1)) / benefits.length
          benefits.forEach((b, i) => {
            const x = MX + i * (bw + gap)
            renderCard(s, x, by2, bw, bh, pal(i), b.title ?? '', b.desc ?? '', { P, PL, INK, MUT, WHITE, FONT })
          })
        }
        drawFooter(s, slide, slideNumber, { PD, MUT, WHITE, FONT })
        continue
      }

      if (variant === 'metrics') {
        // Baris 3 big-number + kartu TANTANGAN vs DAMPAK
        const metrics = (bc.metrics ?? []).slice(0, 3)
        const my = yTop
        const mh = 1.35
        if (metrics.length) {
          const gap = 0.23
          const mw = (CW - gap * (metrics.length - 1)) / metrics.length
          metrics.forEach((m, i) => {
            const x = MX + i * (mw + gap)
            s.addShape('roundRect', { x, y: my, w: mw, h: mh, fill: { color: WHITE }, rectRadius: 0.06, line: { type: 'none' } })
            s.addShape('rect', { x, y: my, w: 0.1, h: mh, fill: { color: pal(i) } })
            addText(String(m.big ?? ''), {
              x: x + 0.3, y: my + 0.12, w: mw - 0.5, h: 0.55,
              fontSize: fitPt(String(m.big ?? ''), mw - 0.5, 0.55, 26, 12), color: P, bold: true,
            })
            addText(String(m.desc ?? ''), {
              x: x + 0.3, y: my + 0.72, w: mw - 0.5, h: mh - 0.85,
              fontSize: fitPt(String(m.desc ?? ''), mw - 0.5, mh - 0.85, 11, 8), color: MUT,
              lineSpacingMultiple: 1.25,
            })
          })
        }
        const cy = my + mh + 0.25
        const chh = Math.min(2.7, Math.max(1.4, FOOTER_Y - cy - 0.12))
        const cwid = (CW - 0.23) / 2
        s.addShape('roundRect', { x: MX, y: cy, w: cwid, h: chh, fill: { color: WHITE }, rectRadius: 0.04, line: { type: 'none' } })
        addText(L_CHAL, { x: MX + 0.3, y: cy + 0.18, w: cwid - 0.6, h: 0.28, fontSize: 11, color: P, bold: true, fit: 'none' })
        {
          const fit = clampFit(String(bc.challenge ?? ''), cwid - 0.6, chh - 0.7)
          addText(fit.text, {
            x: MX + 0.3, y: cy + 0.5, w: cwid - 0.6, h: chh - 0.7,
            fontSize: fit.pt, color: MUT, lineSpacingMultiple: 1.35,
          })
        }
        const rx = MX + cwid + 0.23
        s.addShape('roundRect', { x: rx, y: cy, w: cwid, h: chh, fill: { color: PD }, rectRadius: 0.04, line: { type: 'none' } })
        addText(L_IMP, { x: rx + 0.3, y: cy + 0.18, w: cwid - 0.6, h: 0.28, fontSize: 11, color: WHITE, bold: true, fit: 'none' })
        {
          const fit = clampFit(String(bc.impact ?? ''), cwid - 0.6, chh - 0.7)
          addText(fit.text, {
            x: rx + 0.3, y: cy + 0.5, w: cwid - 0.6, h: chh - 0.7,
            fontSize: fit.pt, color: 'F0EAEA', lineSpacingMultiple: 1.35,
          })
        }
        drawFooter(s, slide, slideNumber, { PD, MUT, WHITE, FONT })
        continue
      }

      // variant full: 2 kartu + tabel (kartu memanjang bila tanpa tabel)
      const tbl = bc.table
      const hasTable = !!(tbl?.headers?.length && tbl?.rows?.length)
      const cy = yTop + 0.05
      const chh = hasTable ? 2.18 : Math.max(2.5, FOOTER_Y - 0.15 - cy)
      const cwid = (CW - 0.23) / 2
      s.addShape('roundRect', { x: MX, y: cy, w: cwid, h: chh, fill: { color: WHITE }, rectRadius: 0.04, line: { type: 'none' } })
      addText(L_BG, { x: MX + 0.3, y: cy + 0.18, w: cwid - 0.6, h: 0.28, fontSize: 11, color: P, bold: true, fit: 'none' })
      const bgText = String(bc.background ?? '').trim()
      const probText = String(bc.business_problem ?? '').trim()
      if (probText) {
        const bgH = hasTable ? 0.82 : 1.7
        const probLblY = cy + 0.48 + bgH + 0.1
        const bgFit = clampFit(bgText, cwid - 0.6, bgH)
        addText(bgFit.text, {
          x: MX + 0.3, y: cy + 0.48, w: cwid - 0.6, h: bgH,
          fontSize: bgFit.pt, color: MUT, lineSpacingMultiple: 1.3,
        })
        addText(L_PROB, { x: MX + 0.3, y: probLblY, w: cwid - 0.6, h: 0.28, fontSize: 11, color: PL, bold: true, fit: 'none' })
        const probH = cy + chh - probLblY - 0.46
        const probFit = clampFit(probText, cwid - 0.6, probH)
        addText(probFit.text, {
          x: MX + 0.3, y: probLblY + 0.28, w: cwid - 0.6, h: probH,
          fontSize: probFit.pt, color: MUT, lineSpacingMultiple: 1.3,
        })
      } else {
        const bgFit = clampFit(bgText, cwid - 0.6, chh - 0.68)
        addText(bgFit.text, {
          x: MX + 0.3, y: cy + 0.48, w: cwid - 0.6, h: chh - 0.68,
          fontSize: bgFit.pt, color: MUT, lineSpacingMultiple: 1.35,
        })
      }
      const rx = MX + cwid + 0.23
      s.addShape('roundRect', { x: rx, y: cy, w: cwid, h: chh, fill: { color: PD }, rectRadius: 0.04, line: { type: 'none' } })
      addText(L_SOL, { x: rx + 0.3, y: cy + 0.18, w: cwid - 0.6, h: 0.28, fontSize: 11, color: WHITE, bold: true, fit: 'none' })
      const solParas = (Array.isArray(bc.solution) ? bc.solution : [String(bc.solution ?? '')])
        .map((p) => String(p).trim()).filter(Boolean).slice(0, 4)
      if (solParas.length) {
        const sy = cy + 0.5
        const segH = (chh - 0.66) / solParas.length
        solParas.forEach((p, i) => {
          const fit = clampFit(p, cwid - 0.6, segH - 0.04)
          addText(fit.text, {
            x: rx + 0.3, y: sy + i * segH, w: cwid - 0.6, h: segH - 0.04,
            fontSize: fit.pt, color: 'F0EAEA', lineSpacingMultiple: 1.3,
          })
        })
      }
      if (hasTable) {
        const tt = String(bc.table_title ?? '').trim()
        let ty = cy + chh + 0.14
        if (tt) {
          addText(tt.toUpperCase(), { x: MX, y: ty, w: CW, h: 0.3, fontSize: 11, color: P, bold: true, fit: 'none' })
          ty += 0.34
        }
        const headers = tbl!.headers.map(String)
        const rows = (tbl!.rows ?? []).slice(0, 5).map((r) => r.map(String))
        const ncols = headers.length
        const colW = [CW * 0.2, ...Array(Math.max(ncols - 2, 1)).fill((CW * 0.68) / Math.max(ncols - 2, 1)), CW * 0.12]
        s.addTable(
          [
            headers.map((h, ci) => ({
              text: h,
              options: {
                bold: true, color: WHITE, fill: { color: PD }, fontFace: 'Georgia', fontSize: 11,
                valign: 'middle' as const, align: ci === ncols - 1 ? ('center' as const) : ('left' as const),
              },
            })),
            ...rows.map((r, ri) =>
              Array.from({ length: ncols }, (_, ci) => ({
                text: r[ci] ?? '',
                options: {
                  bold: ci === 0 || ci === ncols - 1,
                  color: ci === 0 ? INK : ci === ncols - 1 ? P : MUT,
                  fill: { color: ri % 2 === 0 ? WHITE : ZEBRA },
                  fontFace: 'Georgia', fontSize: 10.5, valign: 'middle' as const,
                  align: ci === ncols - 1 ? ('center' as const) : ('left' as const),
                },
              })),
            ),
          ],
          { x: MX, y: ty, w: CW, colW, border: { type: 'none' as const }, rowH: 0.34 },
        )
      }
      drawFooter(s, slide, slideNumber, { PD, MUT, WHITE, FONT })
      continue
    }

    // ══════════ CONTENT SLIDES ══════════
    // Gambar hanya bila benar-benar terunduh dan konten tidak penuh.
    const hasImage = !!imageData && layout !== 'stats' && layout !== 'callout'
    const imageLeft = slide.imageSide === 'left'
    const contentX = hasImage && imageLeft ? 5.4 : MX
    const contentW = hasImage ? CW - 4.6 : CW

    let y = 0.55
    if (slide.eyebrow) {
      addText(slide.eyebrow.toUpperCase(), {
        x: contentX, y, w: contentW, h: 0.35, fontSize: 12, color: ACC, bold: true, charSpacing: 2,
        fit: 'none',
      })
      y += 0.42
    }
    addText(slide.title, {
      x: contentX, y, w: contentW, h: 0.85, fontSize: fitPt(slide.title, contentW, 0.85, 28, 16),
      color: INK, bold: true,
    })
    y += 0.95
    // Intro → banner KEY MESSAGE house-style (rounded + bar aksen)
    if (slide.intro?.trim()) {
      s.addShape('roundRect', { x: contentX, y, w: contentW, h: 0.75, fill: { color: BANNER }, rectRadius: 0.04, line: { type: 'none' } })
      s.addShape('rect', { x: contentX, y, w: 0.09, h: 0.75, fill: { color: PL } })
      addText(slide.intro.trim(), {
        x: contentX + 0.3, y: y + 0.04, w: contentW - 0.55, h: 0.67,
        fontSize: fitPt(slide.intro, contentW - 0.55, 0.67, 14, 9), color: INK, bold: true,
        valign: 'middle', lineSpacingMultiple: 1.2,
      })
      y += 0.9
    }
    y += 0.15

    const content = slide.content ?? []

    if (layout === 'callout') {
      const quote =
        typeof content[0] === 'string'
          ? (content[0] as string)
          : ((content[0] as { title?: string })?.title ?? '')
      const qy = Math.max(y, 1.9)
      const qh = 3.4
      s.addShape('roundRect', { x: contentX, y: qy, w: contentW, h: qh, fill: { color: P }, rectRadius: 0.06, line: { type: 'none' } })
      s.addShape('rect', { x: contentX, y: qy, w: 0.12, h: qh, fill: { color: PL } })
      addText(`\u201C${quote}\u201D`, {
        x: contentX + 0.5, y: qy + 0.4, w: contentW - 1, h: qh - 1.3,
        fontSize: fitPt(quote, contentW - 1, qh - 1.3, 24, 12), color: WHITE, italic: true,
        lineSpacingMultiple: 1.3,
      })
      if (slide.intro?.trim())
        addText(`\u2014 ${slide.intro.trim()}`, {
          x: contentX + 0.5, y: qy + qh - 0.65, w: contentW - 1, h: 0.4, fontSize: 13, color: 'D9CFCF',
          bold: true, fit: 'none',
        })
      drawFooter(s, slide, slideNumber, { PD, MUT, WHITE, FONT })
      continue
    }

    if (layout === 'agenda' && typeof content[0] === 'string') {
      const items = content as string[]
      const availH = FOOTER_Y - y
      const gap = 0.12
      const ih = Math.min(0.85, (availH - (items.length - 1) * gap) / Math.max(items.length, 1))
      items.forEach((item, i) => {
        const ry = y + i * (ih + gap)
        s.addShape('roundRect', { x: contentX, y: ry, w: 0.55, h: ih, fill: { color: P }, rectRadius: 0.06, line: { type: 'none' } })
        addText(String(i + 1), {
          x: contentX, y: ry, w: 0.55, h: ih, fontSize: 18, color: WHITE, bold: true,
          align: 'center', valign: 'middle',
        })
        addText(item, {
          x: contentX + 0.8, y: ry, w: contentW - 0.8, h: ih,
          fontSize: fitPt(item, contentW - 0.8, ih, 16, 10), color: INK, valign: 'middle',
        })
      })
      drawFooter(s, slide, slideNumber, { PD, MUT, WHITE, FONT })
      continue
    }

    if (layout === 'cards') {
      const cards = (content as Array<{ title: string; desc?: string }>).slice(0, 3)
      const availH = FOOTER_Y - y - 0.1
      const chh = Math.min(3.6, availH)
      const gap = 0.23
      const cwid = (contentW - gap * Math.max(cards.length - 1, 0)) / Math.max(cards.length, 1)
      cards.forEach((c, i) =>
        renderCard(s, contentX + i * (cwid + gap), y, cwid, chh, pal(i), c.title ?? '', c.desc ?? '', {
          P, PL, INK, MUT, WHITE, FONT,
        }),
      )
      drawFooter(s, slide, slideNumber, { PD, MUT, WHITE, FONT })
      continue
    }

    if (layout === 'stats') {
      const stats = (content as Array<{ big: string; desc: string }>).slice(0, 4)
      const availH = FOOTER_Y - y - 0.1
      const gap = 0.23
      const n = Math.max(stats.length, 1)
      const cwid = (contentW - gap * (n - 1)) / n
      // Kartu SEJAJAR horizontal — tinggi cukup 1.9" (bukan dibagi per kartu!)
      const chh = Math.min(1.9, availH)
      // House-style (ported from fastapi _fill_stats): font seragam per slide,
      // desc PANJANG DIPOTONG (bukan diperkecil tanpa batas) — floor 12pt.
      const descW = cwid - 0.6
      const charsPerLine = Math.max(14, Math.floor(descW / 0.075))
      const maxDescLen = Math.max(...stats.map((st) => (st.desc ?? '').length), 0)
      const estLines = maxDescLen ? Math.max(1, Math.ceil(maxDescLen / charsPerLine)) : 1
      const descPt = estLines <= 1 ? 15 : estLines <= 2 ? 14 : 13
      const maxChars = charsPerLine * 3
      stats.forEach((st, i) => {
        const x = contentX + i * (cwid + gap)
        s.addShape('roundRect', { x, y, w: cwid, h: chh, fill: { color: WHITE }, rectRadius: 0.06, line: { type: 'none' } })
        s.addShape('rect', { x, y, w: 0.12, h: chh, fill: { color: pal(i) } })
        addText(String(st.big ?? ''), {
          x: x + 0.3, y: y + 0.15, w: Math.min(2.4, cwid - 0.6), h: 0.66,
          fontSize: fitPt(String(st.big ?? ''), Math.min(2.4, cwid - 0.6), 0.66, 30, 14), color: P,
          bold: true,
        })
        let d = String(st.desc ?? '')
        if (d.length > maxChars) d = `${d.slice(0, maxChars - 1).trimEnd()}…`
        const fit = clampFit(d, descW, chh - 1.0, descPt, 12)
        addText(fit.text, {
          x: x + 0.3, y: y + 0.9, w: descW, h: chh - 1.05,
          fontSize: fit.pt, color: MUT, lineSpacingMultiple: 1.25,
        })
      })
      drawFooter(s, slide, slideNumber, { PD, MUT, WHITE, FONT })
      continue
    }

    // numbered_list (juga target redirect title_content/rows/timeline/comparison)
    {
      type Row = { title: string; desc: string }
      const rows: Row[] = (content as Array<string | { title?: string; desc?: string; big?: string }>).map(
        (c) => {
          if (typeof c === 'string') return { title: c, desc: '' }
          return { title: String(c.title ?? c.big ?? ''), desc: String(c.desc ?? '') }
        },
      ).filter((r) => r.title || r.desc)
      const n = Math.max(rows.length, 1)
      const availH = FOOTER_Y - y - 0.05
      const gap = 0.12
      const ih = Math.min(1.3, (availH - (n - 1) * gap) / n)
      const maxDescLen = Math.max(...rows.map((r) => r.desc.length), 0)
      const descW = contentW - 1.35
      const estLines = maxDescLen ? Math.max(1, Math.ceil(maxDescLen / (descW / 0.075))) : 1
      const descPt = estLines <= 1 ? 16 : estLines <= 2 ? 14 : 13
      const titlePt = n <= 3 ? 20 : n <= 4 ? 18 : 16
      rows.forEach((r, i) => {
        const ry = y + i * (ih + gap)
        s.addShape('roundRect', { x: contentX, y: ry, w: contentW, h: ih, fill: { color: WHITE }, rectRadius: 0.08, line: { type: 'none' } })
        s.addShape('rect', { x: contentX, y: ry, w: 0.07, h: ih, fill: { color: pal(i) } })
        addText(String(i + 1).padStart(2, '0'), {
          x: contentX + 0.25, y: ry, w: 0.85, h: ih, fontSize: 24, color: P, bold: true,
          valign: 'middle',
        })
        const tx = contentX + 1.2
        const tw = contentW - (tx - contentX) - 0.25
        if (r.desc) {
          addText(r.title, {
            x: tx, y: ry + 0.02, w: tw, h: ih * 0.42,
            fontSize: fitPt(r.title, tw, ih * 0.42, titlePt, 10), color: INK, bold: true,
            valign: 'bottom',
          })
          const descFit = clampFit(r.desc, tw, ih * 0.52, descPt, 10)
          addText(descFit.text, {
            x: tx, y: ry + ih * 0.46, w: tw, h: ih * 0.52,
            fontSize: descFit.pt, color: MUT,
            valign: 'top', lineSpacingMultiple: 1.2,
          })
        } else {
          addText(r.title, {
            x: tx, y: ry + 0.04, w: tw, h: ih - 0.08,
            fontSize: fitPt(r.title, tw, ih - 0.08, titlePt, 10), color: INK, valign: 'middle',
            lineSpacingMultiple: 1.25,
          })
        }
      })
    }

    // Footer house-style: bar tipis + label + nomor
    drawFooter(s, slide, slideNumber, { PD, MUT, WHITE, FONT })

    // Gambar konten (numbered_list/cards) — sisi sesuai imageSide
    if (hasImage) {
      const imgX = imageLeft ? MX : W - MX - 4.4
      s.addImage({
        data: imageData!,
        x: imgX,
        y: 1.2,
        w: 4.4,
        h: 5.4,
        sizing: { type: 'cover', w: 4.4, h: 5.4 },
      })
    }
  }

  const buf = await pptx.write({ outputType: 'arraybuffer' })
  return { bytes: new Uint8Array(buf as ArrayBuffer), imageFailures }
}

/** Kartu house-style: rounded putih + band warna atas + dot + title + desc (seragam, tanpa kartu gelap). */
function renderCard(
  s: import('pptxgenjs').default.Slide,
  x: number,
  y: number,
  w: number,
  h: number,
  colorBand: string,
  title: string,
  desc: string,
  C: { P: string; PL: string; INK: string; MUT: string; WHITE: string; FONT: { fontFace: string } },
): void {
  const FONT = { fontFace: 'Georgia' }
  s.addShape('roundRect', { x, y, w, h, fill: { color: C.WHITE }, rectRadius: 0.06, line: { type: 'none' } })
  s.addShape('roundRect', { x, y, w, h: 0.12, fill: { color: colorBand }, rectRadius: 0.06, line: { type: 'none' } })
  const hasDot = !!title
  if (hasDot) {
    s.addShape('ellipse', { x: x + 0.3, y: y + 0.35, w: 0.5, h: 0.5, fill: { color: colorBand } })
    s.addText((title[0] ?? '').toUpperCase(), {
      ...FONT, x: x + 0.3, y: y + 0.35, w: 0.5, h: 0.5, fontSize: 18, color: C.WHITE, bold: true,
      align: 'center', valign: 'middle', fit: 'none',
    })
    s.addText(title, {
      ...FONT, x: x + 0.95, y: y + 0.4, w: w - 1.2, h: 0.75,
      fontSize: fitPt(title, w - 1.2, 0.75, 17, 9), color: C.INK, bold: true, valign: 'middle',
      fit: 'shrink',
    })
    if (desc) {
      const fit = clampFit(desc, w - 0.6, h - 1.3, 12, 10)
      s.addText(fit.text, {
        ...FONT, x: x + 0.3, y: y + 1.1, w: w - 0.6, h: h - 1.3,
        fontSize: fit.pt, color: C.MUT, valign: 'top',
        lineSpacingMultiple: 1.3, fit: 'shrink',
      })
    }
  } else {
    s.addText(desc, {
      ...FONT, x: x + 0.3, y: y + 0.3, w: w - 0.6, h: h - 0.6,
      fontSize: fitPt(desc, w - 0.6, h - 0.6, 13, 8), color: C.MUT, valign: 'top', fit: 'shrink',
    })
  }
}

/** Footer house-style: bar tipis primaryDark + nomor halaman kanan. */
function drawFooter(
  s: import('pptxgenjs').default.Slide,
  slide: SlideContent,
  slideNumber: number,
  C: { PD: string; MUT: string; WHITE: string; FONT: { fontFace: string } },
): void {
  s.addShape('rect', { x: 0, y: 7.38, w: 13.33, h: 0.12, fill: { color: C.PD } })
  s.addText(String(slideNumber), {
    ...C.FONT, x: 12.4, y: 7.02, w: 0.7, h: 0.3, fontSize: 10, color: C.MUT, align: 'right',
    valign: 'middle',
  })
}
