import ELK from 'elkjs/lib/elk.bundled.js'
import { NODE_SIZE, type FlowDocument, type FlowEdge, type FlowNode, type NodeType } from '../shared/ipc'

const elk = new ELK()

export interface NodeGeometry {
  x: number
  y: number
  w: number
  h: number
  cx: number
  cy: number
}

export function nodeGeometry(node: FlowNode): NodeGeometry {
  const { w, h } = NODE_SIZE[node.type]
  return { x: node.x, y: node.y, w, h, cx: node.x + w / 2, cy: node.y + h / 2 }
}

/** Run elkjs layered layout; returns a new doc with nodes repositioned (topâ†’bottom flow). */
export async function autoLayout(doc: FlowDocument): Promise<FlowDocument> {
  if (doc.nodes.length === 0) return doc
  const children = doc.nodes.map((n) => ({
    id: n.id,
    width: NODE_SIZE[n.type].w,
    height: NODE_SIZE[n.type].h,
    labels: [{ text: n.label || n.type }],
  }))
  const edges = doc.edges.map((e) => ({
    id: e.id,
    sources: [e.from],
    targets: [e.to],
    ...(e.label ? { labels: [{ text: e.label }] } : {}),
  }))
  const graph = {
    id: 'root',
    layoutOptions: {
      'elk.algorithm': 'layered',
      'elk.direction': 'DOWN',
      'elk.spacing.nodeNode': '48',
      'elk.layered.spacing.nodeNodeBetweenLayers': '64',
      'elk.padding': '[top=40,left=40,bottom=40,right=40]',
    },
    children,
    edges,
  }
  try {
    const res = await elk.layout(graph)
    const posById = new Map<string, { x: number; y: number }>()
    for (const c of res.children ?? []) {
      posById.set(c.id, { x: Math.round(c.x ?? 0), y: Math.round(c.y ?? 0) })
    }
    return {
      ...doc,
      nodes: doc.nodes.map((n) => {
        const p = posById.get(n.id)
        return p ? { ...n, x: p.x, y: p.y } : n
      }),
    }
  } catch {
    return doc
  }
}

export const NODE_TYPE_LABEL: Record<NodeType, string> = {
  start: 'Start',
  end: 'End',
  process: 'Process',
  decision: 'Decision',
  io: 'Input / Output',
}

let idCounter = 0
export function nextId(prefix: string): string {
  idCounter += 1
  return `${prefix}-${Date.now().toString(36)}-${idCounter}`
}

/* â”€â”€ Swimlane layout (vertical lanes = process owners) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
export interface LaneLayout {
  lanes: Array<{ id: string; name: string; x: number; width: number }>
  headerH: number
  rowStep: number
  topPad: number
  leftPad: number
  totalW: number
  totalH: number
}

const LANE_MIN_W = 180
const LANE_MAX_W = 280
const HEADER_H = 48
const ROW_STEP = 156
const TOP_PAD = 24
const LEFT_PAD = 24

function estimateLaneWidth(name: string): number {
  const w = (name || '').length * 8 + 40
  return Math.max(LANE_MIN_W, Math.min(LANE_MAX_W, w))
}

/** Longest-path rank per node (cycle-safe). */
function computeRanks(doc: FlowDocument): Map<string, number> {
  const rank = new Map<string, number>()
  const preds = new Map<string, string[]>()
  for (const n of doc.nodes) preds.set(n.id, [])
  for (const e of doc.edges) if (preds.has(e.to)) preds.get(e.to)!.push(e.from)
  const visiting = new Set<string>()
  const resolve = (id: string): number => {
    const cached = rank.get(id)
    if (cached !== undefined) return cached
    if (visiting.has(id)) return 0
    visiting.add(id)
    let r = 0
    for (const p of preds.get(id) ?? []) r = Math.max(r, resolve(p) + 1)
    visiting.delete(id)
    rank.set(id, r)
    return r
  }
  for (const n of doc.nodes) resolve(n.id)
  return rank
}

export function computeLaneLayout(doc: FlowDocument): LaneLayout {
  const lanes = doc.lanes && doc.lanes.length ? doc.lanes : [{ id: '_default', name: 'Flow' }]
  let x = LEFT_PAD
  const laid = lanes.map((l) => {
    const width = estimateLaneWidth(l.name)
    const item = { id: l.id, name: l.name, x, width }
    x += width
    return item
  })
  const ranks = computeRanks(doc)
  const maxRank = ranks.size ? Math.max(...Array.from(ranks.values())) : 0
  const totalW = x + LEFT_PAD
  const totalH = TOP_PAD + HEADER_H + (maxRank + 1) * ROW_STEP + 40
  return { lanes: laid, headerH: HEADER_H, rowStep: ROW_STEP, topPad: TOP_PAD, leftPad: LEFT_PAD, totalW, totalH }
}

/** Position nodes into lane columns (x = lane, y = rank), with per-rank row heights so nothing overlaps. */
export function autoLayoutWithLanes(doc: FlowDocument): FlowDocument {
  if (!doc.lanes || doc.lanes.length === 0) return doc
  const layout = computeLaneLayout(doc)
  const laneById = new Map(layout.lanes.map((l) => [l.id, l]))
  const ranks = computeRanks(doc)
  // Wide enough that edge routing can thread a corridor between stacked same-rank nodes.
  const V_GAP = 56

  const cellMap = new Map<string, { laneId: string; rank: number; ids: string[] }>()
  for (const n of doc.nodes) {
    const laneId = n.lane && laneById.has(n.lane) ? n.lane : layout.lanes[0].id
    const rank = ranks.get(n.id) ?? 0
    const key = `${laneId}:${rank}`
    let cell = cellMap.get(key)
    if (!cell) { cell = { laneId, rank, ids: [] }; cellMap.set(key, cell) }
    cell.ids.push(n.id)
  }
  const rowHeight = new Map<number, number>()
  for (const cell of cellMap.values()) {
    let h = 0
    for (const id of cell.ids) {
      const node = doc.nodes.find((x) => x.id === id)
      if (node) h += NODE_SIZE[node.type].h + V_GAP
    }
    h = Math.max(60, h - V_GAP)
    rowHeight.set(cell.rank, Math.max(rowHeight.get(cell.rank) ?? 0, h))
  }
  const maxRank = ranks.size ? Math.max(...Array.from(ranks.values())) : 0
  const rankY = new Map<number, number>()
  let y = layout.topPad + layout.headerH
  for (let r = 0; r <= maxRank; r++) {
    rankY.set(r, y)
    y += (rowHeight.get(r) ?? 60) + 64
  }
  const offsetInCell = new Map<string, number>()
  const nodes = doc.nodes.map((n) => {
    const laneId = n.lane && laneById.has(n.lane) ? n.lane : layout.lanes[0].id
    const lane = laneById.get(laneId)!
    const nw = NODE_SIZE[n.type].w
    const rank = ranks.get(n.id) ?? 0
    const key = `${laneId}:${rank}`
    const off = offsetInCell.get(key) ?? 0
    offsetInCell.set(key, off + NODE_SIZE[n.type].h + V_GAP)
    const x = lane.x + (lane.width - nw) / 2
    return { ...n, lane: laneId, x: Math.round(x), y: Math.round((rankY.get(rank) ?? 0) + off) }
  })
  return { ...doc, nodes }
}
/* â”€â”€ Edge routing with per-node ports (decisions branch bottom + side) â”€ */
export interface EdgeGeom {
  id: string
  points: number[]
  labelX: number
  labelY: number
  labelW: number
}

/**
 * Compute orthogonal edge geometry. For a decision with >=2 outgoing edges, the edge
 * whose target sits most directly below exits from the BOTTOM; the others exit from the
 * LEFT/RIGHT side nearest their target â€” so branches are visually distinct and short.
 */
/** Anchor point on a node's boundary for a given side. */
type Side = 'top' | 'bottom' | 'left' | 'right'

/** Boundary anchor, skew-aware for the parallelogram (io) so lines glue to the slanted edge.
 *  `off` slides the anchor along the side so several edges can share one side without piling up. */
function anchor(g: NodeGeometry, side: Side, type: NodeType, off = 0): { x: number; y: number } {
  const k = type === 'io' ? 18 : 0
  if (side === 'top') return { x: g.x + (k + g.w) / 2 + off, y: g.y }
  if (side === 'bottom') return { x: g.x + (g.w - k) / 2 + off, y: g.y + g.h }
  if (side === 'left') return { x: g.x + k / 2, y: g.cy + off }
  return { x: g.x + g.w - k / 2, y: g.cy + off }
}

function outward(side: Side): { x: number; y: number } {
  if (side === 'top') return { x: 0, y: -1 }
  if (side === 'bottom') return { x: 0, y: 1 }
  if (side === 'left') return { x: -1, y: 0 }
  return { x: 1, y: 0 }
}

/** Simple orthogonal elbow (fallback when A* finds no path). */
function ortho(sx: number, sy: number, sp: Side, tx: number, ty: number, tp: Side): number[] {
  if ((sp === 'bottom' && tp === 'top') || (sp === 'top' && tp === 'bottom')) {
    const my = (sy + ty) / 2
    return [sx, sy, sx, my, tx, my, tx, ty]
  }
  if ((sp === 'right' && tp === 'left') || (sp === 'left' && tp === 'right')) {
    const mx = (sx + tx) / 2
    return [sx, sy, mx, sy, mx, ty, tx, ty]
  }
  if ((sp === 'left' || sp === 'right') && tp === 'top') return [sx, sy, tx, sy, tx, ty]
  if (sp === 'bottom' && (tp === 'left' || tp === 'right')) return [sx, sy, sx, ty, tx, ty]
  return [sx, sy, tx, ty]
}

const CELL = 16
const PAD = 44
// Obstacle inflation for the routing grid (must stay under half the tightest inter-node gap).
const INF = 4
// Must exceed cleanArrowEnd's 20px minimum so extending the arrow never skews the stub.
const STUB = 24
// Shared corridor length below which two edges are allowed to ride the same line.
const MIN_SPREAD = 24

interface Grid { ox: number; oy: number; cols: number; rows: number; blocked: Set<string> }

function buildGrid(doc: FlowDocument): Grid | null {
  if (doc.nodes.length === 0) return null
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity
  for (const n of doc.nodes) {
    const g = nodeGeometry(n)
    minX = Math.min(minX, g.x); minY = Math.min(minY, g.y)
    maxX = Math.max(maxX, g.x + g.w); maxY = Math.max(maxY, g.y + g.h)
  }
  minX -= PAD; minY -= PAD; maxX += PAD; maxY += PAD
  const cols = Math.ceil((maxX - minX) / CELL)
  const rows = Math.ceil((maxY - minY) / CELL)
  const blocked = new Set<string>()
  for (const n of doc.nodes) {
    const g = nodeGeometry(n)
    // Small inflation + floor rounding: a cell is blocked only when it truly overlaps the
    // inflated rect. Heavier inflation (or ceil rounding) seals the 28px channels between
    // stacked nodes, so A* cannot reach the goal and edges fall back to straight pierces.
    const x0 = Math.floor((g.x - INF - minX) / CELL), y0 = Math.floor((g.y - INF - minY) / CELL)
    const x1 = Math.floor((g.x + g.w + INF - minX) / CELL), y1 = Math.floor((g.y + g.h + INF - minY) / CELL)
    for (let cx = x0; cx <= x1; cx++) for (let cy = y0; cy <= y1; cy++) blocked.add(cx + ',' + cy)
  }
  return { ox: minX, oy: minY, cols, rows, blocked }
}

function cellCenter(grid: Grid, cx: number, cy: number): { x: number; y: number } {
  return { x: grid.ox + (cx + 0.5) * CELL, y: grid.oy + (cy + 0.5) * CELL }
}

/** A* on a 4-dir grid with a turn penalty (fewer bends), avoiding node cells. */
function routeAstar(grid: Grid, s: { x: number; y: number }, g: { x: number; y: number }): number[] | null {
  const toCell = (pt: { x: number; y: number }) => ({
    cx: Math.max(0, Math.min(grid.cols - 1, Math.floor((pt.x - grid.ox) / CELL))),
    cy: Math.max(0, Math.min(grid.rows - 1, Math.floor((pt.y - grid.oy) / CELL))),
  })
  const sc = toCell(s), gc = toCell(g)
  const allow = new Set<string>()
  for (const [ux, uy] of [[0, 0], [1, 0], [-1, 0], [0, 1], [0, -1]]) {
    allow.add(sc.cx + ux + ',' + (sc.cy + uy))
    allow.add(gc.cx + ux + ',' + (gc.cy + uy))
  }
  const dirs = [[1, 0], [-1, 0], [0, 1], [0, -1]]
  const key = (cx: number, cy: number, d: number) => cx + ',' + cy + ',' + d
  const gScore = new Map<string, number>()
  const parent = new Map<string, string>()
  const open: { cx: number; cy: number; dir: number; f: number }[] = []
  const h = (cx: number, cy: number) => Math.abs(cx - gc.cx) + Math.abs(cy - gc.cy)
  gScore.set(key(sc.cx, sc.cy, -1), 0)
  open.push({ cx: sc.cx, cy: sc.cy, dir: -1, f: h(sc.cx, sc.cy) })
  let guard = 0
  while (open.length && guard++ < 120000) {
    let bi = 0
    for (let i = 1; i < open.length; i++) if (open[i].f < open[bi].f) bi = i
    const cur = open.splice(bi, 1)[0]
    if (cur.cx === gc.cx && cur.cy === gc.cy) {
      const cells: { cx: number; cy: number }[] = []
      let k = key(cur.cx, cur.cy, cur.dir)
      const seen = new Set<string>()
      while (k && !seen.has(k)) {
        seen.add(k)
        const parts = k.split(',')
        cells.push({ cx: +parts[0], cy: +parts[1] })
        const par = parent.get(k)
        if (par === undefined) break
        k = par
      }
      cells.reverse()
      const pts: number[] = []
      for (const c of cells) { const cc = cellCenter(grid, c.cx, c.cy); pts.push(cc.x, cc.y) }
      return pts
    }
    for (let d = 0; d < 4; d++) {
      const ncx = cur.cx + dirs[d][0], ncy = cur.cy + dirs[d][1]
      if (ncx < 0 || ncy < 0 || ncx >= grid.cols || ncy >= grid.rows) continue
      const isGoal = ncx === gc.cx && ncy === gc.cy
      if (grid.blocked.has(ncx + ',' + ncy) && !isGoal && !allow.has(ncx + ',' + ncy)) continue
      const turn = cur.dir !== -1 && cur.dir !== d ? 8 : 0
      const ng = (gScore.get(key(cur.cx, cur.cy, cur.dir)) ?? 0) + 1 + turn
      const nk = key(ncx, ncy, d)
      if (gScore.has(nk) && (gScore.get(nk) as number) <= ng) continue
      gScore.set(nk, ng)
      parent.set(nk, key(cur.cx, cur.cy, cur.dir))
      open.push({ cx: ncx, cy: ncy, dir: d, f: ng + h(ncx, ncy) })
    }
  }
  return null
}

function simplify(pts: number[]): number[] {
  if (pts.length <= 4) return pts
  const out: number[] = [pts[0], pts[1]]
  for (let i = 2; i < pts.length - 2; i += 2) {
    const px = out[out.length - 2], py = out[out.length - 1]
    const cx = pts[i], cy = pts[i + 1]
    const nx = pts[i + 2], ny = pts[i + 3]
    const collinear = (Math.abs(px - cx) < 1 && Math.abs(cx - nx) < 1) || (Math.abs(py - cy) < 1 && Math.abs(cy - ny) < 1)
    if (!collinear) out.push(cx, cy)
  }
  out.push(pts[pts.length - 2], pts[pts.length - 1])
  return out
}

/** Force every segment axis-aligned: split any diagonal into an L (corner). */
function makeOrthogonal(pts: number[], sp: Side): number[] {
  if (pts.length < 4) return pts
  const out: number[] = [pts[0], pts[1]]
  let lastAxis: 'h' | 'v' = sp === 'top' || sp === 'bottom' ? 'v' : 'h'
  for (let i = 2; i < pts.length; i += 2) {
    const px = out[out.length - 2], py = out[out.length - 1]
    const cx = pts[i], cy = pts[i + 1]
    const ddx = Math.abs(px - cx), ddy = Math.abs(py - cy)
    if (ddx > 0.5 && ddy > 0.5) {
      if (lastAxis === 'v') out.push(px, cy)
      else out.push(cx, py)
      lastAxis = lastAxis === 'v' ? 'h' : 'v'
    } else {
      lastAxis = ddx > ddy ? 'h' : 'v'
    }
    out.push(cx, cy)
  }
  return out
}

/**
/**
 * Orthogonal edge routing glued to node boundaries, with A* obstacle avoidance so
 * lines do not cut through intermediate symbols. Decisions send the most-vertical
 * branch from the bottom and the others from the nearest side.
 */
/** Nearest side of g to og's center. */
function nearestSide(g: NodeGeometry, og: NodeGeometry, type: NodeType): Side {
  let best: Side = 'bottom'
  let bestD = Infinity
  for (const s of ['top', 'bottom', 'left', 'right'] as Side[]) {
    const a = anchor(g, s, type)
    const d = Math.abs(a.x - og.cx) + Math.abs(a.y - og.cy)
    if (d < bestD) { bestD = d; best = s }
  }
  return best
}

/** Side facing the other node's center (fall back to nearest anchor when overlapping bands). */
function dominantSide(g: NodeGeometry, og: NodeGeometry, type: NodeType): Side {
  if (og.cy < g.y) return 'top'
  if (og.cy > g.y + g.h) return 'bottom'
  if (og.cx < g.x) return 'left'
  if (og.cx > g.x + g.w) return 'right'
  return nearestSide(g, og, type)
}

const PORT_SPREAD = 26

/** Assign each edge a source/target side; edges sharing a side get spread-out anchor offsets
 *  so arrows never stack on one point and never get pushed to a geometrically wrong side. */
function assignPorts(
  doc: FlowDocument,
  byId: Map<string, FlowNode>,
): Map<string, { sp: Side; tp: Side; so: number; to: number }> {
  const result = new Map<string, { sp: Side; tp: Side; so: number; to: number }>()
  const nodeEdges = new Map<string, FlowEdge[]>()
  for (const e of doc.edges) {
    if (!nodeEdges.has(e.from)) nodeEdges.set(e.from, [])
    if (!nodeEdges.has(e.to)) nodeEdges.set(e.to, [])
    nodeEdges.get(e.from)!.push(e)
    nodeEdges.get(e.to)!.push(e)
  }
  for (const [nodeId, edgesAt] of nodeEdges) {
    const node = byId.get(nodeId)
    if (!node) continue
    const g = nodeGeometry(node)
    const bySide = new Map<Side, { edge: FlowEdge; other: NodeGeometry }[]>()
    for (const e of edgesAt) {
      const other = byId.get(e.from === nodeId ? e.to : e.from)
      if (!other) continue
      const side = dominantSide(g, nodeGeometry(other), node.type)
      if (!bySide.has(side)) bySide.set(side, [])
      bySide.get(side)!.push({ edge: e, other: nodeGeometry(other) })
    }
    // Flowchart convention for decisions: the branch continuing straight down leaves the
    // bottom vertex; the other branches leave from the LEFT/RIGHT side facing their target.
    if (node.type === 'decision') {
      const bottom = bySide.get('bottom')
      if (bottom && bottom.length > 1) {
        bottom.sort((a, b) => Math.abs(a.other.cx - g.cx) - Math.abs(b.other.cx - g.cx))
        for (const item of bottom.splice(1)) {
          let side: Side
          if (item.other.cx < g.cx - 8) side = 'left'
          else if (item.other.cx > g.cx + 8) side = 'right'
          else side = (bySide.get('left')?.length ?? 0) <= (bySide.get('right')?.length ?? 0) ? 'left' : 'right'
          if (!bySide.has(side)) bySide.set(side, [])
          bySide.get(side)!.push(item)
        }
      }
    }
    for (const [side, group] of bySide) {
      const horizontal = side === 'left' || side === 'right'
      group.sort((a, b) => (horizontal ? a.other.cy - b.other.cy : a.other.cx - b.other.cx))
      const span = horizontal ? g.h : g.w
      const maxHalf = Math.max(0, span / 2 - 20)
      const step =
        group.length > 1 ? Math.min(PORT_SPREAD, (maxHalf * 2) / (group.length - 1)) : 0
      group.forEach((item, i) => {
        const off = Math.round((i - (group.length - 1) / 2) * step)
        const cur = result.get(item.edge.id) ?? {
          sp: 'bottom' as Side,
          tp: 'top' as Side,
          so: 0,
          to: 0,
        }
        if (item.edge.from === nodeId) {
          cur.sp = side
          cur.so = off
        } else {
          cur.tp = side
          cur.to = off
        }
        result.set(item.edge.id, cur)
      })
    }
  }
  return result
}


/** True if any polyline segment crosses a node rect (excluding source/target). */
function crossesAnyNode(points: number[], doc: FlowDocument, srcId: string, tgtId: string): boolean {
  for (const n of doc.nodes) {
    if (n.id === srcId || n.id === tgtId) continue
    const g = nodeGeometry(n)
    for (let i = 0; i < points.length - 2; i += 2) {
      const x1 = points[i], y1 = points[i + 1], x2 = points[i + 2], y2 = points[i + 3]
      if (y1 === y2) {
        if (y1 > g.y && y1 < g.y + g.h && Math.max(x1, x2) > g.x && Math.min(x1, x2) < g.x + g.w) return true
      } else {
        if (x1 > g.x && x1 < g.x + g.w && Math.max(y1, y2) > g.y && Math.min(y1, y2) < g.y + g.h) return true
      }
    }
  }
  return false
}

/** Ensure the final segment (into the target) is long enough so the arrowhead is clean. */
function cleanArrowEnd(points: number[]): number[] {
  if (points.length < 6) return points
  const n = points.length
  const x1 = points[n - 4], y1 = points[n - 3], x2 = points[n - 2], y2 = points[n - 1]
  const len = Math.hypot(x2 - x1, y2 - y1)
  if (len < 20 && len > 0.001) {
    const ux = (x2 - x1) / len, uy = (y2 - y1) / len
    points[n - 4] = x2 - ux * 20
    points[n - 3] = y2 - uy * 20
  }
  return points
}

/** Snap A* cell-center points back onto the anchor axis. The grid's cell centers sit up to
 *  half a cell away from node centerlines, which would otherwise turn every straight edge
 *  into an 8px staircase; detour points farther than that are left untouched.
 *  Additionally, the departure and arrival runs are pulled flush onto the stub axes so
 *  lines leave/enter nodes dead straight (no last-moment 3-8px jog before the arrow). */
function snapAxes(
  pts: number[],
  sa: { x: number; y: number },
  ta: { x: number; y: number },
  ss?: { x: number; y: number },
  tt?: { x: number; y: number },
  sp?: Side,
  tp?: Side,
): number[] {
  const T = CELL / 2 + 1
  const out: number[] = [pts[0], pts[1]]
  for (let i = 2; i < pts.length - 2; i += 2) {
    let x = pts[i]
    let y = pts[i + 1]
    if (Math.abs(x - sa.x) <= T && Math.abs(x - ta.x) <= T) x = sa.x
    if (Math.abs(y - sa.y) <= T && Math.abs(y - ta.y) <= T) y = sa.y
    out.push(x, y)
  }
  out.push(pts[pts.length - 2], pts[pts.length - 1])
  // pts layout for the stub variant: [sa, ss, ...path, tt, ta] (path starts at index 4)
  if (ss && tt && sp && tp && out.length >= 10) {
    const exitHorizontal = sp === 'left' || sp === 'right'
    for (let i = 4; i <= out.length - 6; i += 2) {
      if (exitHorizontal ? Math.abs(out[i + 1] - ss.y) <= T : Math.abs(out[i] - ss.x) <= T) {
        if (exitHorizontal) out[i + 1] = ss.y
        else out[i] = ss.x
      } else break
    }
    const entryHorizontal = tp === 'left' || tp === 'right'
    for (let i = out.length - 6; i >= 4; i -= 2) {
      if (entryHorizontal ? Math.abs(out[i + 1] - tt.y) <= T : Math.abs(out[i] - tt.x) <= T) {
        if (entryHorizontal) out[i + 1] = tt.y
        else out[i] = tt.x
      } else break
    }
  }
  return simplify(out)
}

function segHitsRect(x1: number, y1: number, x2: number, y2: number, g: NodeGeometry): boolean {
  if (y1 === y2) return y1 > g.y && y1 < g.y + g.h && Math.max(x1, x2) > g.x && Math.min(x1, x2) < g.x + g.w
  if (x1 === x2) return x1 > g.x && x1 < g.x + g.w && Math.max(y1, y2) > g.y && Math.min(y1, y2) < g.y + g.h
  return false
}

/** Last resort when A* cannot thread a corridor (tight stacked rows): wrap the offending
 *  segment around the node it cuts, with perpendicular stubs so arrows stay clean.
 *  A node that gets hit twice is bypassed on the opposite side; a third hit means the
 *  geometry is unwinnable and we stop rather than oscillate. */
function detourAround(pts: number[], doc: FlowDocument, srcId: string, tgtId: string): number[] {
  const M = 18
  const others = doc.nodes.filter((n) => n.id !== srcId && n.id !== tgtId)
  const hits = new Map<string, number>()
  for (let iter = 0; iter < 10; iter++) {
    let hit: { i: number; g: NodeGeometry; n: FlowNode } | null = null
    outer: for (let i = 0; i < pts.length - 2; i += 2) {
      for (const n of others) {
        const g = nodeGeometry(n)
        if (segHitsRect(pts[i], pts[i + 1], pts[i + 2], pts[i + 3], g)) {
          hit = { i, g, n }
          break outer
        }
      }
    }
    if (!hit) return pts
    const { i, g, n } = hit
    const seen = (hits.get(n.id) ?? 0) + 1
    hits.set(n.id, seen)
    if (seen > 2) return pts
    const flip = seen % 2 === 0
    const x1 = pts[i], y1 = pts[i + 1], x2 = pts[i + 2], y2 = pts[i + 3]
    let repl: number[]
    if (y1 === y2) {
      const preferTop = Math.abs(y1 - (g.y - M)) <= Math.abs(y1 - (g.y + g.h + M))
      const goTop = preferTop !== flip
      const wy = goTop ? g.y - M : g.y + g.h + M
      repl = [x1, wy, x2, wy, x2, y2]
    } else {
      const preferLeft = Math.abs(x1 - (g.x - M)) <= Math.abs(x1 - (g.x + g.w + M))
      const goLeft = preferLeft !== flip
      const wx = goLeft ? g.x - M : g.x + g.w + M
      const s = y2 >= y1 ? M : -M
      let hy1 = y1 + s
      let hy2 = y2 - s
      // keep the wrap horizontals outside the node's vertical span so they never graze its border
      if (hy1 >= g.y && hy1 <= g.y + g.h) hy1 = y2 >= y1 ? g.y - M : g.y + g.h + M
      if (hy2 >= g.y && hy2 <= g.y + g.h) hy2 = y2 >= y1 ? g.y + g.h + M : g.y - M
      repl = [x1, hy1, wx, hy1, wx, hy2, x2, hy2, x2, y2]
    }
    pts = [...pts.slice(0, i + 2), ...repl, ...pts.slice(i + 4)]
  }
  return pts
}

/** First collinear run two polylines share (longer than MIN_SPREAD; shorter stubs at a
 *  shared lane center are unavoidable and visually harmless, so we leave them alone). */
function sharedRun(a: number[], b: number[]): { dir: 'v' | 'h'; fixed: number } | null {
  for (let i = 0; i < a.length - 2; i += 2) {
    for (let k = 0; k < b.length - 2; k += 2) {
      const [x1, y1, x2, y2] = [a[i], a[i + 1], a[i + 2], a[i + 3]]
      const [x3, y3, x4, y4] = [b[k], b[k + 1], b[k + 2], b[k + 3]]
      if (x1 === x2 && x3 === x4 && x1 === x3) {
        const lo = Math.max(Math.min(y1, y2), Math.min(y3, y4))
        const hi = Math.min(Math.max(y1, y2), Math.max(y3, y4))
        if (hi - lo > MIN_SPREAD) return { dir: 'v', fixed: x1 }
      }
      if (y1 === y2 && y3 === y4 && y1 === y3) {
        const lo = Math.max(Math.min(x1, x2), Math.min(x3, x4))
        const hi = Math.min(Math.max(x1, x2), Math.max(x3, x4))
        if (hi - lo > MIN_SPREAD) return { dir: 'h', fixed: y1 }
      }
    }
  }
  return null
}

/** Move every interior point of a collinear run sideways (endpoints stay glued to nodes). */
function shiftRun(pts: number[], dir: 'v' | 'h', fixed: number, d: number): void {
  for (let i = 2; i < pts.length - 2; i += 2) {
    if (dir === 'v' ? pts[i] === fixed : pts[i + 1] === fixed) {
      if (dir === 'v') pts[i] = fixed + d
      else pts[i + 1] = fixed + d
    }
  }
}

/** Separate edges that ride the same corridor so each line stays individually readable. */
function spreadOverlaps(all: EdgeGeom[], doc: FlowDocument, sps: Map<string, Side>): void {
  const edgeById = new Map(doc.edges.map((e) => [e.id, e]))
  for (let iter = 0; iter < 4; iter++) {
    let moved = false
    for (let i = 0; i < all.length; i++) {
      for (let j = i + 1; j < all.length; j++) {
        for (let round = 0; round < 3; round++) {
          const ov = sharedRun(all[i].points, all[j].points)
          if (!ov) break
          const e = edgeById.get(all[j].id)
          if (!e) break
          const saved = all[j].points.slice()
          let done = false
          for (const d of [12, -12]) {
            shiftRun(all[j].points, ov.dir, ov.fixed, d)
            all[j].points = simplify(
              makeOrthogonal(all[j].points, sps.get(all[j].id) ?? 'bottom'),
            )
            if (!crossesAnyNode(all[j].points, doc, e.from, e.to)) { done = true; break }
            all[j].points = saved.slice()
          }
          if (done) moved = true
          else {
            all[j].points = saved
            break
          }
        }
      }
    }
    if (!moved) return
  }
}

function polylineLen(pts: number[]): number {
  let len = 0
  for (let i = 0; i < pts.length - 2; i += 2) len += Math.abs(pts[i + 2] - pts[i]) + Math.abs(pts[i + 3] - pts[i + 1])
  return len
}

function pointAtDistance(pts: number[], d: number): { x: number; y: number } {
  let rem = d
  for (let i = 0; i < pts.length - 2; i += 2) {
    const x1 = pts[i], y1 = pts[i + 1], x2 = pts[i + 2], y2 = pts[i + 3]
    const len = Math.abs(x2 - x1) + Math.abs(y2 - y1)
    if (rem <= len || i === pts.length - 4) {
      const t = len === 0 ? 0 : Math.min(1, rem / len)
      return { x: x1 + (x2 - x1) * t, y: y1 + (y2 - y1) * t }
    }
    rem -= len
  }
  return { x: pts[pts.length - 2], y: pts[pts.length - 1] }
}

/** Place the edge label near its source (where branch meaning matters), sliding along the
 *  polyline until the label box clears every node rect instead of landing on a diamond. */
function freeLabelSpot(pts: number[], labelW: number, doc: FlowDocument): { x: number; y: number } {
  const total = polylineLen(pts)
  const fallback = pointAtDistance(pts, Math.min(34, total))
  for (let d = 26; d <= total; d += 12) {
    const p = pointAtDistance(pts, d)
    const bx = p.x - labelW / 2 - 4
    const by = p.y - 13
    const bw = labelW + 8
    const bh = 26
    let ok = true
    for (const n of doc.nodes) {
      const g = nodeGeometry(n)
      if (bx < g.x + g.w && bx + bw > g.x && by < g.y + g.h && by + bh > g.y) { ok = false; break }
    }
    if (ok) return p
  }
  return fallback
}

/**
 * Orthogonal edge routing glued to node boundaries with A* obstacle avoidance.
 * Ports are assigned per node (facing side, spread when shared); the first/last segments
 * are forced perpendicular to the node surface; paths are snapped back onto anchor axes;
 * anything that still cuts a node is wrapped around it, and shared corridors are separated.
 */
export function computeEdgeGeometry(doc: FlowDocument): EdgeGeom[] {
  const byId = new Map(doc.nodes.map((n) => [n.id, n]))
  const ports = assignPorts(doc, byId)
  const grid = buildGrid(doc)
  const geoms: EdgeGeom[] = []
  const sps = new Map<string, Side>()
  for (const e of doc.edges) {
    const a = byId.get(e.from)
    const b = byId.get(e.to)
    if (!a || !b) continue
    const ga = nodeGeometry(a)
    const gb = nodeGeometry(b)
    const p = ports.get(e.id) ?? { sp: 'bottom' as Side, tp: 'top' as Side, so: 0, to: 0 }
    const sp = p.sp
    const tp = p.tp
    const sa = anchor(ga, sp, a.type, p.so)
    const ta = anchor(gb, tp, b.type, p.to)
    const sOff = outward(sp), tOff = outward(tp)
    // Short perpendicular stubs glued to each node surface keep arrows leaving/entering
    // at right angles; the "force perpendicular ends" coordinate hack used to skew these.
    const ss = { x: sa.x + sOff.x * STUB, y: sa.y + sOff.y * STUB }
    const tt = { x: ta.x + tOff.x * STUB, y: ta.y + tOff.y * STUB }
    let points: number[]
    if (grid) {
      const s = { x: sa.x + sOff.x * CELL, y: sa.y + sOff.y * CELL }
      const t = { x: ta.x + tOff.x * CELL, y: ta.y + tOff.y * CELL }
      const path = routeAstar(grid, s, t)
      points = path
        ? snapAxes([sa.x, sa.y, ss.x, ss.y, ...path, tt.x, tt.y, ta.x, ta.y], sa, ta, ss, tt, sp, tp)
        : [sa.x, sa.y, ...ortho(ss.x, ss.y, sp, tt.x, tt.y, tp), ta.x, ta.y]
    } else {
      points = [sa.x, sa.y, ...ortho(ss.x, ss.y, sp, tt.x, tt.y, tp), ta.x, ta.y]
    }
    points = snapAxes(points, sa, ta)
    points = makeOrthogonal(points, sp)
    points = simplify(points)
    points = makeOrthogonal(points, sp)
    // if the path still cuts through a node, retry A* with a wider start/goal offset
    if (grid && crossesAnyNode(points, doc, a.id, b.id)) {
      const s2 = { x: sa.x + sOff.x * CELL * 2, y: sa.y + sOff.y * CELL * 2 }
      const t2 = { x: ta.x + tOff.x * CELL * 2, y: ta.y + tOff.y * CELL * 2 }
      const path2 = routeAstar(grid, s2, t2)
      if (path2) {
        points = snapAxes([sa.x, sa.y, ss.x, ss.y, ...path2, tt.x, tt.y, ta.x, ta.y], sa, ta, ss, tt, sp, tp)
        points = makeOrthogonal(points, sp)
        points = simplify(points)
      }
    }
    points = simplify(detourAround(points, doc, a.id, b.id))
    points = cleanArrowEnd(points)
    // normalize BEFORE overlap spreading so spreadOverlaps sees the true orthogonal shape
    // (cleanArrowEnd/detour can leave diagonals that would later unfold onto a shared line)
    points = simplify(makeOrthogonal(points, sp))
    sps.set(e.id, sp)
    geoms.push({ id: e.id, points, labelX: 0, labelY: 0, labelW: 0 })
  }
  spreadOverlaps(geoms, doc, sps)
  for (const g of geoms) {
    // final pass: shifts can leave tiny diagonals; force everything axis-aligned
    g.points = simplify(makeOrthogonal(g.points, sps.get(g.id) ?? 'bottom'))
  }
  for (const g of geoms) {
    const e = doc.edges.find((x) => x.id === g.id)
    g.labelW = Math.max(36, Math.round((e?.label ?? '').length * 6.8 + 16))
    const lp = freeLabelSpot(g.points, g.labelW, doc)
    g.labelX = lp.x
    g.labelY = lp.y
  }
  return geoms
}
export function validateFlow(doc: FlowDocument): string[] {
  const problems: string[] = []
  const { nodes, edges } = doc
  if (nodes.length === 0) return ['Diagram masih kosong.']
  const byId = new Map(nodes.map((n) => [n.id, n]))
  const out = new Map<string, FlowEdge[]>()
  const indeg = new Map<string, number>()
  for (const n of nodes) { out.set(n.id, []); indeg.set(n.id, 0) }
  for (const e of edges) {
    if (!byId.has(e.from) || !byId.has(e.to)) { problems.push(`Edge ${e.id} menunjuk node yang tidak ada.`); continue }
    out.get(e.from)!.push(e)
    indeg.set(e.to, (indeg.get(e.to) ?? 0) + 1)
  }
  const starts = nodes.filter((n) => n.type === 'start')
  const ends = nodes.filter((n) => n.type === 'end')
  if (starts.length === 0) problems.push("Butuh tepat satu node 'start'.")
  if (starts.length > 1) problems.push(`Hanya boleh satu 'start', ada ${starts.length}.`)
  if (ends.length === 0) problems.push("Butuh minimal satu node 'end'.")
  for (const n of nodes) {
    const outs = out.get(n.id) ?? []
    if (n.type === 'decision') {
      if (outs.length < 2) problems.push(`Decision "${n.label}" (${n.id}) harus punya minimal 2 keluaran (mis. Ya/Tidak).`)
      for (const e of outs) if (!e.label || !e.label.trim()) problems.push(`Cabang decision "${n.label}" ke "${byId.get(e.to)?.label ?? e.to}" wajib diberi label kondisi (mis. Ya/Tidak).`)
    } else if (n.type === 'process' || n.type === 'io') {
      if (outs.length > 1) problems.push(`${n.type === 'io' ? 'I/O' : 'Proses'} "${n.label}" (${n.id}) hanya boleh punya maksimal 1 keluaran (percabangan hanya dari decision).`)
    } else if (n.type === 'start') {
      if (outs.length !== 1) problems.push(`Start "${n.label}" harus punya tepat 1 keluaran (sekarang ${outs.length}).`)
    } else if (n.type === 'end') {
      if (outs.length > 0) problems.push(`End "${n.label}" tidak boleh punya keluaran.`)
    }
  }
  // reachability from start
  if (starts.length >= 1) {
    const seen = new Set<string>()
    const queue = starts.map((s) => s.id)
    while (queue.length) {
      const id = queue.shift()!
      if (seen.has(id)) continue
      seen.add(id)
      for (const e of out.get(id) ?? []) if (!seen.has(e.to)) queue.push(e.to)
    }
    for (const n of nodes) if (!seen.has(n.id)) problems.push(`Node "${n.label}" (${n.id}) tidak terjangkau dari start (gantung/salah sambung).`)
  }
  return problems
}