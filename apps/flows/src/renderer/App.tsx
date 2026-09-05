import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { Stage, Layer, Rect, Ellipse, Line, Text, Arrow, Group } from 'react-konva'
import type Konva from 'konva'
import {
  NODE_SIZE,
  emptyFlow,
  type AiSettings,
  type FlowDocument,
  type FlowNode,
  type NodeType,
} from '../shared/ipc'
import { autoLayout, autoLayoutWithLanes, computeEdgeGeometry, computeLaneLayout, nextId, nodeGeometry } from './flow-model'
import { createFlowsSkill } from './ai/flows-skill'
import { AiPanel } from './ai/AiPanel'

type Mode = 'select' | 'connect'

const NODE_COLORS: Record<NodeType, string> = {
  start: '#2e7d32',
  end: '#c62828',
  process: '#1a73e8',
  decision: '#f9a825',
  io: '#6a1b9a',
}

function shapeFor(node: FlowNode, selected: boolean) {
  const { w, h } = NODE_SIZE[node.type]
  const stroke = selected ? '#111' : NODE_COLORS[node.type]
  const strokeWidth = selected ? 2.5 : 1.5
  const fill = '#ffffff'
  if (node.type === 'start' || node.type === 'end') {
    return <Ellipse x={w / 2} y={h / 2} radiusX={w / 2} radiusY={h / 2} fill={fill} stroke={stroke} strokeWidth={strokeWidth} />
  }
  if (node.type === 'decision') {
    const pts = [w / 2, 0, w, h / 2, w / 2, h, 0, h / 2]
    return <Line points={pts} closed fill={fill} stroke={stroke} strokeWidth={strokeWidth} />
  }
  if (node.type === 'io') {
    const k = 18
    const pts = [k, 0, w, 0, w - k, h, 0, h]
    return <Line points={pts} closed fill={fill} stroke={stroke} strokeWidth={strokeWidth} />
  }
  return <Rect width={w} height={h} fill={fill} stroke={stroke} strokeWidth={strokeWidth} cornerRadius={6} />
}

export function App() {
  const [doc, setDoc] = useState<FlowDocument>(emptyFlow())
  const [selected, setSelected] = useState<string | null>(null)
  const [mode, setMode] = useState<Mode>('select')
  const [connectFrom, setConnectFrom] = useState<string | null>(null)
  const [size, setSize] = useState({ w: 900, h: 700 })
  const [zoom, setZoom] = useState(1)
  const [lang, setLang] = useState('en')
  const docRef = useRef(doc)
  docRef.current = doc
  const historyRef = useRef<FlowDocument[]>([])
  const futureRef = useRef<FlowDocument[]>([])
  const commit = useCallback((next: FlowDocument) => {
    historyRef.current.push(docRef.current)
    if (historyRef.current.length > 100) historyRef.current.shift()
    futureRef.current = []
    setDoc(next)
  }, [])
  const undo = useCallback(() => {
    if (historyRef.current.length === 0) return
    futureRef.current.push(docRef.current)
    setDoc(historyRef.current.pop() as FlowDocument)
  }, [])
  const redo = useCallback(() => {
    if (futureRef.current.length === 0) return
    historyRef.current.push(docRef.current)
    setDoc(futureRef.current.pop() as FlowDocument)
  }, [])
  const settingsRef = useRef<AiSettings>({ provider: 'prova', providers: {} } as unknown as AiSettings)
  const wrapRef = useRef<HTMLDivElement | null>(null)

  useEffect(() => {
    void window.flowsApi.getAiSettings().then((s) => (settingsRef.current = s)).catch(() => {})
    void window.flowsApi.getLanguage().then((l) => setLang(l)).catch(() => {})
  }, [])

  useEffect(() => {
    const measure = (): void => {
      const el = wrapRef.current
      if (el) setSize({ w: el.clientWidth, h: el.clientHeight })
    }
    measure()
    window.addEventListener('resize', measure)
    return () => window.removeEventListener('resize', measure)
  }, [])

  useEffect(() => {
    void window.flowsApi.consumePendingOpen().then((res) => {
      if (res?.doc) { historyRef.current = []; futureRef.current = []; setDoc(res.doc) }
    })
  }, [])

  const access = useMemo(
    () => ({ getDoc: () => docRef.current, setDoc: (d: FlowDocument) => commit(d) }),
    [],
  )
  const skill = useMemo(() => createFlowsSkill(access), [access])

  const addNode = useCallback(
    async (type: NodeType) => {
      const node: FlowNode = { id: nextId('n'), type, label: type, x: 0, y: 0 }
      const next = { ...docRef.current, nodes: [...docRef.current.nodes, node] }
      commit(next.lanes && next.lanes.length ? autoLayoutWithLanes(next) : await autoLayout(next))
    },
    [],
  )

  const onNodeClick = (id: string): void => {
    if (mode === 'connect') {
      if (!connectFrom) {
        setConnectFrom(id)
      } else if (connectFrom !== id) {
        commit({
          ...docRef.current,
          edges: [...docRef.current.edges, { id: nextId('e'), from: connectFrom, to: id }],
        })
        setConnectFrom(null)
      }
    } else {
      setSelected(id)
    }
  }

  const renameSelected = (): void => {
    if (!selected) return
    const node = docRef.current.nodes.find((n) => n.id === selected)
    if (!node) return
    const label = window.prompt('Label', node.label)
    if (label != null)
      commit({
        ...docRef.current,
        nodes: docRef.current.nodes.map((n) => (n.id === selected ? { ...n, label } : n)),
      })
  }

  const deleteSelected = useCallback((): void => {
    if (!selected) return
    commit({
      ...docRef.current,
      nodes: docRef.current.nodes.filter((n) => n.id !== selected),
      edges: docRef.current.edges.filter((e) => e.from !== selected && e.to !== selected),
    })
    setSelected(null)
  }, [selected])

  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      const mod = e.ctrlKey || e.metaKey
      const tgt = (e.target as HTMLElement)?.tagName
      if (mod && (e.key === 'z' || e.key === 'Z' || e.key === 'y' || e.key === 'Y') && tgt !== 'TEXTAREA' && tgt !== 'INPUT') {
        e.preventDefault()
        if (e.key === 'y' || e.key === 'Y') redo()
        else if (e.shiftKey) redo()
        else undo()
        return
      }
      if ((e.key === 'Delete' || e.key === 'Backspace') && selected) {
        const tag = (e.target as HTMLElement)?.tagName
        if (tag !== 'TEXTAREA' && tag !== 'INPUT') {
          e.preventDefault()
          deleteSelected()
        }
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [selected, deleteSelected, undo, redo])

  const runLayout = async (): Promise<void> =>
    commit(
      docRef.current.lanes && docRef.current.lanes.length
        ? autoLayoutWithLanes(docRef.current)
        : await autoLayout(docRef.current),
    )

  const doSave = async (): Promise<void> => {
    await window.flowsApi.saveFlow(docRef.current)
  }
  const doOpen = async (): Promise<void> => {
    const res = await window.flowsApi.openFlow()
    if (res?.doc) { historyRef.current = []; futureRef.current = []; setDoc(res.doc) }
  }
  const doNew = async (): Promise<void> => {
    // open a fresh tab in the shell instead of wiping the current canvas
    void window.flowsApi.newFlow()
  }
  const doExportPng = (): void => {
    const stage = stageRef.current
    if (!stage) return
    const url = stage.toDataURL({ pixelRatio: 2 })
    void window.flowsApi.exportPng(url)
  }

  const stageRef = useRef<Konva.Stage | null>(null)

  // Stage sized to content so long flows trigger scrollbars on the wrapper
  const laneLayout = doc.lanes && doc.lanes.length ? computeLaneLayout(doc) : null
  const PAD = 80
  const contentW = laneLayout ? laneLayout.totalW : doc.nodes.reduce((m, n) => Math.max(m, n.x + NODE_SIZE[n.type].w), 0) + PAD
  const contentH = laneLayout ? laneLayout.totalH : doc.nodes.reduce((m, n) => Math.max(m, n.y + NODE_SIZE[n.type].h), 0) + PAD
  const edgeGeoms = computeEdgeGeometry(doc)
  const stageW = Math.max(size.w, contentW)
  const stageH = Math.max(size.h, contentH)
  const fitZoom = (): number => {
    const el = wrapRef.current
    if (!el || contentW <= 0 || contentH <= 0) return 1
    const z = Math.min((el.clientWidth - 24) / contentW, (el.clientHeight - 24) / contentH, 1)
    return Math.max(0.4, Math.round(z * 10) / 10)
  }

  return (
    <div className="flows-app" style={{ flexDirection: 'column' }}>
      <div className="flows-toolbar">
        <span className="title">{doc.title || (lang === 'id' ? 'Flow Tanpa Judul' : 'Untitled Flow')}</span>
        <button onClick={() => void addNode('start')}>+ Start</button>
        <button onClick={() => void addNode('process')}>+ Process</button>
        <button onClick={() => void addNode('decision')}>+ Decision</button>
        <button onClick={() => void addNode('io')}>+ I/O</button>
        <button onClick={() => void addNode('end')}>+ End</button>
        <button
          onClick={() => setMode(mode === 'connect' ? 'select' : 'connect')}
          style={{ background: mode === 'connect' ? 'var(--accent-soft)' : undefined }}
        >
          {mode === 'connect' ? 'Connecting…' : 'Connect'}
        </button>
        <button onClick={() => void runLayout()}>Auto layout</button>
        <button onClick={renameSelected}>Rename</button>
        <button onClick={deleteSelected}>Delete</button>
        <button onClick={undo} aria-label="Undo">Undo</button>
        <button onClick={redo} aria-label="Redo">Redo</button>
        <span className="spacer" />
        <button onClick={() => setZoom((z) => Math.max(0.4, Math.round((z - 0.1) * 10) / 10))} aria-label="Zoom out">
          −
        </button>
        <button className="zoom-label" onClick={() => setZoom(1)} aria-label="Reset zoom" title="Reset zoom">
          {Math.round(zoom * 100)}%
        </button>
        <button onClick={() => setZoom((z) => Math.min(2, Math.round((z + 0.1) * 10) / 10))} aria-label="Zoom in">
          +
        </button>
        <button onClick={() => setZoom(fitZoom())} aria-label="Fit to window">
          Fit
        </button>
        <span className="spacer" />
        <button onClick={() => void doNew()}>New</button>
        <button onClick={() => void doOpen()}>Open</button>
        <button onClick={() => void doSave()}>Save</button>
        <button onClick={doExportPng}>PNG</button>
      </div>
      <div className="flows-main">
        <AiPanel skill={skill} getSettings={() => settingsRef.current} />
        <div className="flows-canvas-wrap" ref={wrapRef}>
          <Stage
            ref={stageRef}
            width={Math.round(stageW * zoom)}
            height={Math.round(stageH * zoom)}
            onMouseDown={(e) => {
              if (e.target === e.target.getStage()) setSelected(null)
            }}
          >
            <Layer scaleX={zoom} scaleY={zoom}>
              {laneLayout && (
                <>
                  {laneLayout.lanes.map((l, i) => (
                    <Group key={`lane-${l.id}`}>
                      <Rect x={l.x} y={0} width={l.width} height={stageH} fill={i % 2 === 0 ? '#f6f8fb' : '#ffffff'} />
                      <Line points={[l.x, 0, l.x, stageH]} stroke="#d5dae2" strokeWidth={1} />
                      <Line points={[l.x + l.width, 0, l.x + l.width, stageH]} stroke="#d5dae2" strokeWidth={1} />
                      <Rect x={l.x} y={0} width={l.width} height={laneLayout.headerH} fill="#1f2937" />
                      <Text
                        text={l.name}
                        x={l.x + 8}
                        y={0}
                        width={l.width - 16}
                        height={laneLayout.headerH}
                        align="center"
                        verticalAlign="middle"
                        wrap="word"
                        fontSize={13}
                        fontStyle="bold"
                        fill="#ffffff"
                      />
                    </Group>
                  ))}
                </>
              )}
              {edgeGeoms.map((g) => (
                <Arrow
                  key={g.id}
                  points={g.points}
                  stroke="#667085"
                  fill="#667085"
                  strokeWidth={1.6}
                  lineJoin="round"
                  pointerLength={9}
                  pointerWidth={9}
                />
              ))}
              {edgeGeoms.map((g) => {
                const e = doc.edges.find((x) => x.id === g.id)
                if (!e?.label) return null
                return (
                  <Group key={g.id + '-lbl'} listening={false}>
                    <Rect
                      x={g.labelX - g.labelW / 2 - 4}
                      y={g.labelY - 11}
                      width={g.labelW + 8}
                      height={22}
                      cornerRadius={4}
                      fill="#ffffff"
                      opacity={0.92}
                    />
                    <Text
                      text={e.label}
                      x={g.labelX - g.labelW / 2}
                      y={g.labelY - 11}
                      width={g.labelW}
                      align="center"
                      verticalAlign="middle"
                      height={22}
                      fontSize={12}
                      fontStyle="bold"
                      fill="#3a4252"
                    />
                  </Group>
                )
              })}
              {doc.nodes.map((n) => {
                const g = nodeGeometry(n)
                return (
                  <Group
                    key={n.id}
                    x={n.x}
                    y={n.y}
                    draggable={mode === 'select'}
                    onClick={() => onNodeClick(n.id)}
                    onTap={() => onNodeClick(n.id)}
                    onDragStart={() => {
                      historyRef.current.push(docRef.current)
                      if (historyRef.current.length > 100) historyRef.current.shift()
                      futureRef.current = []
                    }}
                    onDragMove={(e) => {
                      const node = e.target as Konva.Node
                      setDoc({
                        ...docRef.current,
                        nodes: docRef.current.nodes.map((m) =>
                          m.id === n.id ? { ...m, x: node.x(), y: node.y() } : m,
                        ),
                      })
                    }}
                    onDragEnd={(e) => {
                      const node = e.target as Konva.Node
                      setDoc({
                        ...docRef.current,
                        nodes: docRef.current.nodes.map((m) =>
                          m.id === n.id ? { ...m, x: Math.round(node.x()), y: Math.round(node.y()) } : m,
                        ),
                      })
                    }}
                  >
                    {shapeFor(n, selected === n.id)}
                    <Text
                      text={n.label}
                      x={n.type === 'io' ? 24 : n.type === 'decision' ? 16 : 6}
                      width={n.type === 'io' ? g.w - 48 : n.type === 'decision' ? g.w - 32 : g.w - 12}
                      align="center"
                      verticalAlign="middle"
                      height={g.h}
                      fontSize={n.type === 'io' || n.type === 'decision' ? 12 : 13}
                      fill="#1f2328"
                      wrap="word"
                      listening={false}
                    />
                  </Group>
                )
              })}
            </Layer>
          </Stage>
        </div>
      </div>
    </div>
  )
}
