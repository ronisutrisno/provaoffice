import { describe, expect, it } from 'vitest'
import { NODE_SIZE, type FlowDocument } from '../src/shared/ipc'
import { computeEdgeGeometry, validateFlow } from '../src/renderer/flow-model'

/** "Alur Rekrutmen Kandidat" as saved from the app (positions included). */
const doc: FlowDocument = {
  version: 1,
  title: 'Alur Rekrutmen Kandidat',
  lanes: [
    { id: 'l1', name: 'Pelamar' },
    { id: 'l2', name: 'HR / Recruiter' },
    { id: 'l3', name: 'User (Pimpinan)' },
  ],
  nodes: [
    { id: 'n1', type: 'start', label: 'Mulai: Posisi dibuka', x: 234, y: 72, lane: 'l2' },
    { id: 'n2', type: 'io', label: 'Pelamar kirim CV & lamaran', x: 34, y: 196, lane: 'l1' },
    { id: 'n3', type: 'process', label: 'Screening administrasi', x: 214, y: 332, lane: 'l2' },
    { id: 'n4', type: 'decision', label: 'Lolos seleksi administrasi?', x: 214, y: 468, lane: 'l2' },
    { id: 'n5', type: 'process', label: 'Tes & wawancara HR', x: 214, y: 628, lane: 'l2' },
    { id: 'n6', type: 'process', label: 'Wawancara user', x: 394, y: 764, lane: 'l3' },
    { id: 'n7', type: 'decision', label: 'Disetujui user?', x: 394, y: 900, lane: 'l3' },
    { id: 'n8', type: 'process', label: 'Pengajuan & approval offering', x: 214, y: 1060, lane: 'l2' },
    { id: 'n9', type: 'decision', label: 'Kandidat menerima offer?', x: 214, y: 1196, lane: 'l2' },
    { id: 'n10', type: 'process', label: 'Pembuatan kontrak kerja', x: 214, y: 1356, lane: 'l2' },
    { id: 'n11', type: 'process', label: 'Onboarding & hari pertama kerja', x: 214, y: 1592, lane: 'l2' },
    { id: 'n12', type: 'process', label: 'Kirim surat penolakan & arsipkan kandidat', x: 214, y: 1456, lane: 'l2' },
    { id: 'n13', type: 'end', label: 'Selesai', x: 234, y: 1728, lane: 'l2' },
  ],
  edges: [
    { id: 'e1', from: 'n1', to: 'n2', label: 'Publish lowongan' },
    { id: 'e2', from: 'n2', to: 'n3' },
    { id: 'e3', from: 'n3', to: 'n4' },
    { id: 'e4', from: 'n4', to: 'n5', label: 'Ya' },
    { id: 'e5', from: 'n4', to: 'n12', label: 'Tidak' },
    { id: 'e6', from: 'n5', to: 'n6', label: 'Lolos tes' },
    { id: 'e7', from: 'n6', to: 'n7' },
    { id: 'e8', from: 'n7', to: 'n8', label: 'Ya' },
    { id: 'e9', from: 'n7', to: 'n12', label: 'Tidak' },
    { id: 'e10', from: 'n8', to: 'n9', label: 'Offer dikirim' },
    { id: 'e11', from: 'n9', to: 'n10', label: 'Ya' },
    { id: 'e12', from: 'n9', to: 'n12', label: 'Tidak' },
    { id: 'e13', from: 'n10', to: 'n11', label: 'Kontrak ditandatangani' },
    { id: 'e14', from: 'n11', to: 'n13' },
    { id: 'e15', from: 'n12', to: 'n13' },
  ],
}

const byId = new Map(doc.nodes.map((n) => [n.id, n]))
const rect = (id: string) => {
  const n = byId.get(id)!
  const s = NODE_SIZE[n.type]
  return { x: n.x, y: n.y, w: s.w, h: s.h }
}
const geoms = computeEdgeGeometry(doc)
const geom = (id: string) => geoms.find((g) => g.id === id)!

function segHitsRect(x1: number, y1: number, x2: number, y2: number, g: { x: number; y: number; w: number; h: number }): boolean {
  if (y1 === y2) return y1 > g.y && y1 < g.y + g.h && Math.max(x1, x2) > g.x && Math.min(x1, x2) < g.x + g.w
  if (x1 === x2) return x1 > g.x && x1 < g.x + g.w && Math.max(y1, y2) > g.y && Math.min(y1, y2) < g.y + g.h
  return false
}

describe('validateFlow', () => {
  it('accepts the saved recruitment flow', () => {
    expect(validateFlow(doc)).toEqual([])
  })
})

describe('edge routing', () => {
  it('renders aligned vertical edges as single straight segments (no grid staircase)', () => {
    expect(geom('e3').points).toEqual([294, 404, 294, 468])
    expect(geom('e10').points).toEqual([294, 1132, 294, 1196])
  })

  it('never cuts through a node that is not an endpoint', () => {
    const hits: string[] = []
    for (const g of geoms) {
      const e = doc.edges.find((x) => x.id === g.id)!
      for (const n of doc.nodes) {
        if (n.id === e.from || n.id === e.to) continue
        for (let i = 0; i < g.points.length - 2; i += 2) {
          if (segHitsRect(g.points[i], g.points[i + 1], g.points[i + 2], g.points[i + 3], rect(n.id))) {
            hits.push(`${g.id} crosses ${n.id}`)
            break
          }
        }
      }
    }
    expect(hits).toEqual([])
  })

  it('does not let two edges ride the same corridor', () => {
    const overlaps: string[] = []
    // short shared stubs (<= MIN_SPREAD) at a common lane center are acceptable; only
    // long shared runs confuse the reader
    const T = 24
    for (let i = 0; i < geoms.length; i++) {
      for (let j = i + 1; j < geoms.length; j++) {
        const a = geoms[i].points
        const b = geoms[j].points
        for (let x = 0; x < a.length - 2; x += 2) {
          for (let y = 0; y < b.length - 2; y += 2) {
            if (a[x] === a[x + 2] && b[y] === b[y + 2] && a[x] === b[y]) {
              const lo = Math.max(Math.min(a[x + 1], a[x + 3]), Math.min(b[y + 1], b[y + 3]))
              const hi = Math.min(Math.max(a[x + 1], a[x + 3]), Math.max(b[y + 1], b[y + 3]))
              if (hi - lo > T) overlaps.push(`${geoms[i].id}/${geoms[j].id} @x=${a[x]}`)
            }
            if (a[x + 1] === a[x + 3] && b[y + 1] === b[y + 3] && a[x + 1] === b[y + 1]) {
              const lo = Math.max(Math.min(a[x], a[x + 2]), Math.min(b[y], b[y + 2]))
              const hi = Math.min(Math.max(a[x], a[x + 2]), Math.max(b[y], b[y + 2]))
              if (hi - lo > T) overlaps.push(`${geoms[i].id}/${geoms[j].id} @y=${a[x + 1]}`)
            }
          }
        }
      }
    }
    expect(overlaps).toEqual([])
  })

  it('enters the end node from above, not from a U-turn below it', () => {
    const p = geom('e15').points
    const n = p.length
    expect(p[n - 3]).toBeLessThan(rect('n13').y)
  })

  it('leaves decisions with one branch from the bottom and the other from a side', () => {
    const leavesDown = (id: string) => {
      const p = geom(id).points
      return p[0] === p[2] // first segment vertical => bottom vertex exit
    }
    const leavesSide = (id: string) => {
      const p = geom(id).points
      return p[1] === p[3] // first segment horizontal => left/right vertex exit
    }
    // n4: e4 (Ya) down, e5 (Tidak) sideways; n7: e8 down, e9 sideways; n9: e11 down, e12 sideways
    expect(leavesDown('e4') && leavesSide('e5')).toBe(true)
    expect(leavesDown('e8') && leavesSide('e9')).toBe(true)
    expect(leavesDown('e11') && leavesSide('e12')).toBe(true)
  })

  it('has no micro-jog segments (grid staircase leftovers)', () => {
    const jogs: string[] = []
    for (const g of geoms) {
      for (let i = 0; i < g.points.length - 2; i += 2) {
        const len =
          Math.abs(g.points[i + 2] - g.points[i]) + Math.abs(g.points[i + 3] - g.points[i + 1])
        if (len < 8) jogs.push(`${g.id} seg@${i} len=${len}`)
      }
    }
    expect(jogs).toEqual([])
  })
})

describe('edge labels', () => {
  it('keep the label box clear of every node', () => {
    const bad: string[] = []
    for (const g of geoms) {
      const e = doc.edges.find((x) => x.id === g.id)
      if (!e?.label) continue
      const box = { x: g.labelX - g.labelW / 2 - 4, y: g.labelY - 13, w: g.labelW + 8, h: 26 }
      for (const n of doc.nodes) {
        const r = rect(n.id)
        if (box.x < r.x + r.w && box.x + box.w > r.x && box.y < r.y + r.h && box.y + box.h > r.y) {
          bad.push(`"${e.label}" (${g.id}) overlaps ${n.id}`)
        }
      }
    }
    expect(bad).toEqual([])
  })
})
