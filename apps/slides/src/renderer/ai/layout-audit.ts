import type {
  GroupRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
} from '@prova/pptx-render'

/**
 * Deterministic layout audit (modeled on the Google Slides add-in review_google_slides_addin geometry-only checks):
 * pure geometric computation, no LLM calls, no screenshots. Checks four kinds of problems:
 *  1. Elements extending past the canvas
 *  2. Pairwise overlap of content elements (text-text / text-image/media)
 *  3. Text overflowing its text box (uses the render layer's already-laid-out text.contentHeight — exact, not estimated)
 *  4. Low contrast: text nearly invisible against its reconstructed background (solid fills under it)
 * Results are appended to layout tools' return values so the AI "sees" the real post-edit state (write → verify → fix loop).
 */

interface AuditEntry {
  id: string
  type: string
  x: number
  y: number
  w: number
  h: number
  hasText: boolean
  preview: string
  /** Pixels by which the text content height exceeds the box height (only meaningful when >0) */
  overflowPx: number
}

const PREVIEW_MAX = 18

function textPreview(node: ShapeRenderNode): string {
  const t = (node.text?.lines ?? [])
    .map((l) => l.runs.map((r) => r.text).join(''))
    .join(' ')
    .trim()
  return t.length > PREVIEW_MAX ? `${t.slice(0, PREVIEW_MAX)}…` : t
}

/** Collect the top-level nodes that take part in the audit (skip master/layout decoration; a group counts as one box). */
function collectEntries(nodes: RenderNode[]): AuditEntry[] {
  const out: AuditEntry[] = []
  for (const n of nodes) {
    if (n.decoration) continue
    const { x, y, w, h } = n.box
    let hasText = false
    let preview = ''
    let overflowPx = 0
    if (n.type === 'shape' || n.type === 'text') {
      const sn = n as ShapeRenderNode
      preview = textPreview(sn)
      hasText = preview.length > 0
      if (sn.text && hasText) {
        const inner = h - sn.text.insets.t - sn.text.insets.b
        overflowPx = Math.round(sn.text.contentHeight - inner)
      }
    } else if (n.type === 'group') {
      // If any child in the group has text, treat it as text content for overlap detection
      hasText = groupHasText(n as GroupRenderNode)
      preview = '(group)'
    }
    out.push({ id: n.sourceId, type: n.type, x, y, w, h, hasText, preview, overflowPx })
  }
  return out
}

function groupHasText(g: GroupRenderNode): boolean {
  for (const c of g.children) {
    if (c.type === 'group') {
      if (groupHasText(c as GroupRenderNode)) return true
    } else if (c.type === 'shape' || c.type === 'text') {
      const t = (c as ShapeRenderNode).text?.lines ?? []
      if (t.some((l) => l.runs.some((r) => r.text.trim()))) return true
    }
  }
  return false
}

const MEDIA_TYPES = new Set(['picture', 'table', 'chart', 'placeholder-chip'])

/** Whether it's a content element (participates in overlap detection): has text, or is a picture/table/chart. */
function isContent(e: AuditEntry): boolean {
  return e.hasText || MEDIA_TYPES.has(e.type)
}

function label(e: AuditEntry): string {
  return e.preview && e.preview !== '(group)' ? `${e.id}"${e.preview}"` : `${e.id}(${e.type})`
}

const EDGE_TOLERANCE_PX = 8
const OVERFLOW_TOLERANCE_PX = 4
/** Threshold for overlap area as a fraction of the smaller element's area */
const OVERLAP_RATIO = 0.12
/** Absolute overlap area floor (px²), filtering out noise like touching trims */
const OVERLAP_MIN_AREA = 400
/** Background color blocks (≥70% of canvas area) don't participate in overlap detection */
const BACKGROUND_AREA_RATIO = 0.7
const MAX_ISSUES = 12
/** WCAG contrast ratio below which text is effectively invisible (white on a
 *  light tint lands ~1.0–1.2; decorative accents on light sit ~1.7–3.0) */
const CONTRAST_MIN = 1.5

type Rgba = [number, number, number, number]

function parseRenderColor(c: string | undefined | null): Rgba | null {
  if (!c) return null
  const s = c.trim()
  // 8-digit #RRGGBBAA (pptx-engine encodes a:alpha into the hex — color.ts)
  const hex8 = /^#?([0-9a-f]{8})$/i.exec(s)
  if (hex8) {
    const n = parseInt(hex8[1]!.slice(0, 6), 16)
    const a = parseInt(hex8[1]!.slice(6, 8), 16) / 255
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255, a]
  }
  const hex = /^#?([0-9a-f]{6})$/i.exec(s)
  if (hex) {
    const n = parseInt(hex[1]!, 16)
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255, 1]
  }
  const rgb = /rgba?\(([^)]+)\)/i.exec(s)
  if (rgb) {
    const parts = (rgb[1] ?? '').split(',').map((p) => parseFloat(p.trim()))
    if (parts.length >= 3 && parts.every((p) => Number.isFinite(p)))
      return [parts[0]!, parts[1]!, parts[2]!, parts.length >= 4 ? parts[3]! : 1]
  }
  return null
}

function srgbLuminance(r: number, g: number, b: number): number {
  const f = (v: number) => {
    const s = v / 255
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4)
  }
  return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)
}

function contrastRatio(a: Rgba, b: Rgba): number {
  const la = srgbLuminance(a[0], a[1], a[2])
  const lb = srgbLuminance(b[0], b[1], b[2])
  const hi = Math.max(la, lb)
  const lo = Math.min(la, lb)
  return (hi + 0.05) / (lo + 0.05)
}

/** Composite fg (with alpha) over an opaque bg. */
function blendOver(fg: Rgba, bg: Rgba): Rgba {
  const t = Math.min(1, Math.max(0, fg[3]))
  return [fg[0] * t + bg[0] * (1 - t), fg[1] * t + bg[1] * (1 - t), fg[2] * t + bg[2] * (1 - t), 1]
}

/**
 * Contrast check: for every text node, reconstruct its effective background by
 * compositing solid fills beneath it (z-order) over the slide background, then
 * flag runs whose WCAG contrast ratio is below CONTRAST_MIN. Image/gradient
 * backgrounds are skipped (unknowable without pixels).
 */
function checkContrast(slide: RenderSlide): string[] {
  const issues: string[] = []
  const base = parseRenderColor(
    slide.background.kind === 'solid' ? slide.background.color : undefined,
  ) ?? [255, 255, 255, 1]
  slide.nodes.forEach((node, idx) => {
    if (node.decoration) return
    if (node.type !== 'shape' && node.type !== 'text') return
    const sn = node as ShapeRenderNode
    const layout = sn.text
    if (!layout) return
    const runs = layout.lines.flatMap((l) => l.runs.filter((r) => r.text.trim() && !r.isBullet))
    if (runs.length === 0) return
    const cx = node.box.x + node.box.w / 2
    const cy = node.box.y + node.box.h / 2
    let bg = base
    let unknown = false
    for (let j = 0; j < idx; j++) {
      const under = slide.nodes[j]!
      if (under.decoration || under.type !== 'shape') continue
      const uf = (under as ShapeRenderNode).fill
      if (uf.kind === 'none' || uf.kind === 'image') continue
      if (uf.kind === 'gradient') {
        if (
          cx >= under.box.x && cx <= under.box.x + under.box.w &&
          cy >= under.box.y && cy <= under.box.y + under.box.h
        ) unknown = true
        continue
      }
      const c = parseRenderColor(uf.color)
      if (!c) continue
      if (cx >= under.box.x && cx <= under.box.x + under.box.w && cy >= under.box.y && cy <= under.box.y + under.box.h)
        bg = blendOver(c, bg)
    }
    const own = sn.fill
    if (own.kind === 'solid') {
      const c = parseRenderColor(own.color)
      if (c) bg = blendOver(c, bg)
    } else if (own.kind === 'gradient') {
      unknown = true
    }
    if (unknown) return
    let worst = Infinity
    let worstColor = ''
    for (const run of runs) {
      const c = parseRenderColor(run.color)
      if (!c) continue
      const solid = blendOver(c, bg)
      const ratio = contrastRatio(solid, bg)
      if (ratio < worst) { worst = ratio; worstColor = run.color ?? '' }
    }
    if (worst < CONTRAST_MIN) {
      const preview = textPreview(sn)
      issues.push(
        `Low contrast: text ${sn.sourceId}${preview ? ` "${preview}"` : ''} (${worstColor}) has ratio ${worst.toFixed(2)}:1 against its background — nearly invisible. Use dark text on light fills, or white text only on dark fills`,
      )
    }
  })
  return issues
}

/**
 * Audit one page's layout and return the list of problems (empty array = pass).
 */
export function auditSlideLayout(slide: RenderSlide): string[] {
  const entries = collectEntries(slide.nodes)
  const issues: string[] = []
  const W = slide.widthPx
  const H = slide.heightPx

  // 1. Out of bounds
  for (const e of entries) {
    const parts: string[] = []
    if (e.x < -EDGE_TOLERANCE_PX) parts.push(`${Math.round(-e.x)}px past the left edge`)
    if (e.y < -EDGE_TOLERANCE_PX) parts.push(`${Math.round(-e.y)}px past the top edge`)
    if (e.x + e.w > W + EDGE_TOLERANCE_PX)
      parts.push(`${Math.round(e.x + e.w - W)}px past the right edge`)
    if (e.y + e.h > H + EDGE_TOLERANCE_PX)
      parts.push(`${Math.round(e.y + e.h - H)}px past the bottom edge`)
    if (parts.length) issues.push(`Out of bounds: ${label(e)} ${parts.join(', ')}`)
  }

  // 2. Text overflow
  for (const e of entries) {
    if (e.overflowPx > OVERFLOW_TOLERANCE_PX) {
      issues.push(
        `Text overflow: ${label(e)} content exceeds the box height by ${e.overflowPx}px (make the box taller or reduce the font size)`,
      )
    }
  }

  // 3. Pairwise overlap of content elements
  const content = entries.filter((e) => isContent(e) && e.w * e.h < W * H * BACKGROUND_AREA_RATIO)
  for (let i = 0; i < content.length; i++) {
    for (let j = i + 1; j < content.length; j++) {
      const a = content[i]!
      const b = content[j]!
      // Only report text<->text and text<->media; media-on-media (e.g. a chart on an image) is often intentional design, don't report
      if (!a.hasText && !b.hasText) continue
      const ix = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x)
      const iy = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y)
      if (ix <= 0 || iy <= 0) continue
      const inter = ix * iy
      const minArea = Math.min(a.w * a.h, b.w * b.h)
      if (inter < OVERLAP_MIN_AREA || inter < minArea * OVERLAP_RATIO) continue
      issues.push(
        `Overlap: ${label(a)} and ${label(b)} intersect by ${Math.round(ix)}×${Math.round(iy)}px`,
      )
      if (issues.length >= MAX_ISSUES) break
    }
    if (issues.length >= MAX_ISSUES) break
  }

  // 4. Mixed bullet structure: literal "•"/"-" markers mixed with plain lines in one
  //    text box render as an unindented mess (no hanging indent, no visual grouping).
  //    The fix is structural: split into separate bullet paragraphs or separate boxes.
  for (const e of entries) {
    if (!e.hasText || e.type !== 'shape' && e.type !== 'text') continue
    const node = (slide.nodes.find((n) => n.sourceId === e.id) ?? null) as ShapeRenderNode | null
    const lines = (node?.text?.lines ?? []).map((l) => l.runs.map((r) => r.text).join('').trim())
    // Serialized-object leak: an object was stringified somewhere in the pipeline.
    if (lines.some((l) => l.includes('[object Object]'))) {
      issues.push(
        `Broken text: ${label(e)} contains "[object Object]" — pass plain strings (or {title,desc} objects) as content items, never raw objects`,
      )
      continue
    }
    const nonEmpty = lines.filter(Boolean)
    if (nonEmpty.length < 3) continue
    const bulleted = nonEmpty.filter((l) => /^[•·▪◦‣⁃-]\s+/.test(l)).length
    const plain = nonEmpty.length - bulleted
    if (bulleted >= 2 && plain >= 2) {
      issues.push(
        `Mixed structure: ${label(e)} combines ${bulleted} bullet lines with ${plain} plain lines — split bullets into real bullet paragraphs (hanging indent) and headers into their own text box`,
      )
      if (issues.length >= MAX_ISSUES) break
    }
  }

  // 5. Low contrast (text invisible against its reconstructed background)
  for (const issue of checkContrast(slide)) {
    if (issues.length >= MAX_ISSUES) break
    issues.push(issue)
  }

  return issues.slice(0, MAX_ISSUES)
}

/** Format the audit result as trailing text for a tool's return value. */
export function formatAudit(issues: string[], round?: string): string {
  if (issues.length === 0)
    return '\n<layout-audit>✅ Passed: no overlap/out-of-bounds/text overflow/contrast issues.</layout-audit>'
  const head = `\n<layout-audit>⚠️ Found ${issues.length} issue(s):\n`
  const body = issues.map((s) => `- ${s}`).join('\n')
  const tail = round
    ? `\n${round}\n</layout-audit>`
    : "\n→ Immediately write another execute_slide_script to fix these issues (don't stop, don't ask the user, don't declare completion). els reflects the new positions after the last apply; compute from it directly. At most 2 fix rounds; only if still unresolved tell the user honestly.\n</layout-audit>"
  return head + body + tail
}
