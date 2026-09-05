import { useEffect, useRef, useState } from 'react'
import { AgentLoop, type AgentImage, type AgentSkill } from '@prova/agent-core'
import { Markdown } from '@prova/ui'
import { ATTACHMENT_IMAGE_EXTS, type AiSettings, type AttachmentMeta } from '../../shared/ipc'
import { createElectronTransport } from './transport'
import appIcon from '../assets/app-icon.png'

interface ToolActivity {
  name: string
  summary: string
  isError?: boolean | undefined
  done?: boolean | undefined
}

interface ChatMsg {
  role: 'user' | 'assistant'
  text: string
  tools?: ToolActivity[]
  error?: boolean
  startedAt?: number
  endedAt?: number
}

export interface AiPanelProps {
  skill: AgentSkill
  getSettings: () => AiSettings
}

// Build closing-tag patterns without a literal "</" in source (breaks tool-call parsing).
const CLOSE_TC = '<' + '/tool_call>'
const CLOSE_FN = '<' + '/function>'
function stripToolCallXml(text: string): string {
  if (!text.includes('tool_call') && !text.includes('function=')) return text
  return text
    .replace(new RegExp('<tool_call>[\\s\\S]*?' + CLOSE_TC, 'g'), '')
    .replace(new RegExp('<function=[^>]*>[\\s\\S]*?' + CLOSE_FN, 'g'), '')
    .trim()
}

function WorkGroup({
  tools,
  startedAt,
  endedAt,
  now,
}: {
  tools: ToolActivity[]
  startedAt: number
  endedAt?: number
  now: number
}) {
  const running = !endedAt
  const secs = Math.max(0, Math.round(((endedAt ?? now) - startedAt) / 1000))
  const label = running
    ? `Mengerjakanâ€¦ (${tools.length} langkah)`
    : `Selesai (${tools.length} langkah) Â· ${secs}s`
  return (
    <div className="flows-work-group">
      <div className="flows-work-summary">
        {running && (
          <span className="flows-hourglass" aria-hidden="true">
            <span className="flows-spinner" aria-hidden="true" />
          </span>
        )}
        <span>{label}</span>
      </div>
      <div className="flows-work-rows">
        {tools.map((t, j) => (
          <div key={j} className={`flows-work-row${t.done ? ' done' : ''}${t.isError ? ' err' : ''}`}>
            <span className="dot" />
            <span className="flows-work-step">{j + 1}.</span>
            <span>{t.summary}</span>
          </div>
        ))}
      </div>
    </div>
  )
}

export function AiPanel({ skill, getSettings }: AiPanelProps) {
  const [chat, setChat] = useState<ChatMsg[]>([])
  const [input, setInput] = useState('')
  const [busy, setBusy] = useState(false)
  const [now, setNow] = useState(0)
  const [lang, setLang] = useState('en')
  const [attachments, setAttachments] = useState<AttachmentMeta[]>([])
  const [attachNotice, setAttachNotice] = useState<string | null>(null)
  const loopRef = useRef<AgentLoop | null>(null)
  const logRef = useRef<HTMLDivElement | null>(null)
  const skillRef = useRef(skill)
  skillRef.current = skill

  useEffect(() => {
    const loop = new AgentLoop({
      transport: createElectronTransport(getSettings),
      skill: skillRef.current,
      maxTurns: 28,
      events: {
        onText: (text) => {
          setChat((prev) => {
            const last = prev[prev.length - 1]
            if (last && last.role === 'assistant') {
              const copy = prev.slice()
              copy[copy.length - 1] = { ...last, text: stripToolCallXml(text) }
              return copy
            }
            return [...prev, { role: 'assistant', text: stripToolCallXml(text), startedAt: Date.now() }]
          })
        },
        onToolStart: (call) => {
          setChat((prev) => {
            const copy = prev.slice()
            const last = copy[copy.length - 1]
            const act: ToolActivity = { name: call.name, summary: call.name, isError: false, done: false }
            if (last && last.role === 'assistant') {
              copy[copy.length - 1] = { ...last, tools: [...(last.tools ?? []), act] }
            } else {
              copy.push({ role: 'assistant', text: '', tools: [act], startedAt: Date.now() })
            }
            return copy
          })
        },
        onToolExecuted: ({ call, execution }) => {
          setChat((prev) => {
            const copy = prev.slice()
            const last = copy[copy.length - 1]
            if (last && last.role === 'assistant' && last.tools?.length) {
              const tools = last.tools.slice()
              let idx = -1
              for (let k = tools.length - 1; k >= 0; k--) { if (tools[k].name === call.name && !tools[k].done) { idx = k; break } }
              const at = idx >= 0 ? idx : tools.length - 1
              tools[at] = {
                name: call.name,
                summary: execution.summary || call.name,
                isError: execution.isError,
                done: true,
              }
              copy[copy.length - 1] = { ...last, tools }
            }
            return copy
          })
        },
        onDone: () => {
          setBusy(false)
          setNow(Date.now())
          setChat((prev) => {
            const copy = prev.slice()
            const last = copy[copy.length - 1]
            if (last && last.role === 'assistant' && !last.endedAt) copy[copy.length - 1] = { ...last, endedAt: Date.now() }
            return copy
          })
        },
      },
    })
    loopRef.current = loop
    return () => loop.cancel?.()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight
  }, [chat, busy])

  useEffect(() => {
    if (!busy) return
    setNow(Date.now())
    const id = window.setInterval(() => setNow(Date.now()), 1000)
    return () => window.clearInterval(id)
  }, [busy])

  useEffect(() => {
    void window.flowsApi.getLanguage().then((l) => setLang(l)).catch(() => {})
  }, [])

  const pick = (): void => {
    void window.flowsApi.pickAttachments().then((res) => {
      if (!res) return
      if (res.accepted.length) setAttachments((prev) => [...prev, ...res.accepted])
      if (res.rejected.length) {
        setAttachNotice(res.rejected.join('; '))
        window.setTimeout(() => setAttachNotice(null), 5000)
      }
    })
  }

  const removeAttachment = (path: string): void =>
    setAttachments((prev) => prev.filter((a) => a.path !== path))

  const send = async (): Promise<void> => {
    const instruction = input.trim()
    const loop = loopRef.current
    if ((!instruction && attachments.length === 0) || busy || !loop) return
    setBusy(true)
    const atts = attachments
    setAttachments([])
    setInput('')
    let full = instruction || 'Please review the attached document(s).'
    const images: AgentImage[] = []
    const textParts: string[] = []
    for (const a of atts) {
      if (ATTACHMENT_IMAGE_EXTS.has(a.ext)) {
        const r = await window.flowsApi.readAttachmentImage(a.path)
        if (r.ok && r.base64 && r.mime) images.push({ base64: r.base64, mime: r.mime })
      } else {
        const r = await window.flowsApi.readAttachment(a.path)
        if (r.ok && r.kind === 'text' && r.text) textParts.push(`--- ${a.name} ---` + "\n" + r.text)
        else textParts.push(`--- ${a.name} --- (unreadable)`)
      }
    }
    if (textParts.length) full += "\n\n[Attached documents]\n" + textParts.join("\n\n")
    const display = instruction || (atts.length ? atts.map((a) => a.name).join(', ') : '')
    setChat((prev) => [...prev, { role: 'user', text: display }])
    loop.run(full, images)
  }

  const last = chat[chat.length - 1]
  const runningTool = !!last?.tools?.some((t) => !t.done)
  const showTyping = busy && (!last || last.role !== 'assistant' || !runningTool)

  return (
    <div className="flows-ai">
      <div className="flows-ai-head">
        <img src={appIcon} width={18} height={18} alt="" style={{ borderRadius: 3 }} />
        <span>PROVA-AI</span>
      </div>
      <div className="flows-ai-log" ref={logRef}>
        {chat.length === 0 && (
          <div className="flows-msg assistant">
            {lang === 'id'
              ? 'Jelaskan sebuah proses, saya akan membuatkan flowchart-nya. Contoh: "pengajuan cuti: karyawan mengajukan, atasan mereview, jika disetujui HRD mencatat, jika tidak beri notifikasi".'
              : 'Describe a process and I will build the flowchart. Example: "leave approval: request, manager review, if approved HR records it, else notify".'}
          </div>
        )}
        {chat.map((m, i) => (
          <div key={i} className={`flows-msg ${m.role}${m.error ? ' error' : ''}`}>
            {m.tools && m.tools.length > 0 && (
              <WorkGroup
                tools={m.tools}
                startedAt={m.startedAt ?? (now || Date.now())}
                endedAt={m.endedAt}
                now={now || Date.now()}
              />
            )}
            {m.text ? <Markdown text={m.text} /> : null}
          </div>
        ))}
        {showTyping && (
          <div className="flows-msg assistant">
            <div className="flows-typing">
              <span />
              <span />
              <span />
            </div>
          </div>
        )}
      </div>
      {attachNotice && <div className="flows-attach-notice">{attachNotice}</div>}
      {attachments.length > 0 && (
        <div className="flows-attach-chips">
          {attachments.map((a) => (
            <span key={a.path} className="flows-attach-chip" title={a.name}>
              {a.name}
              <button className="flows-attach-x" onClick={() => removeAttachment(a.path)} aria-label="Remove">
                {'Ã—'}
              </button>
            </span>
          ))}
        </div>
      )}
      <div className="flows-ai-input">
        <div className="flows-input-box">
          <textarea
            value={input}
            placeholder="Describe the flow to build or change..."
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault()
                void send()
              }
            }}
          />
          <div className="flows-input-footer">
            <button className="flows-clip-btn" onClick={pick} aria-label="Attach document">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="M21 11.5 12.5 20a5 5 0 0 1-7-7l8.5-8.5a3.3 3.3 0 0 1 4.7 4.7l-8.4 8.4a1.7 1.7 0 0 1-2.4-2.4l7.8-7.8" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
            <span className="flows-input-spacer" />
            <button className="flows-send-btn" onClick={() => void send()} disabled={busy} aria-label="Send">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="M5 12h13M13 6l6 6-6 6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}