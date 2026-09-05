import { Fragment, type CSSProperties, type ReactNode } from 'react'

/**
 * Minimal dependency-free markdown for chat bubbles: paragraphs, ul/ol,
 * headings, **bold**, *italic*, `inline code`, [links](url), fenced code
 * blocks, --- rules, and pipe tables. Tolerates partial (streaming) input —
 * anything unrecognized renders as plain text.
 *
 * Table/code/hr/link styling is applied inline via theme tokens (var(--color-*))
 * so the component renders correctly in every host app without duplicating CSS.
 */

const INLINE_RE = /(`[^`\n]+`|\*\*[^*\n]+?\*\*|\*[^*\n]+?\*|\[[^\]\n]+\]\([^)\n]+\))/g

function renderInline(text: string): ReactNode[] {
  const out: ReactNode[] = []
  let last = 0
  let key = 0
  for (const m of text.matchAll(INLINE_RE)) {
    const i = m.index ?? 0
    if (i > last) out.push(text.slice(last, i))
    const tok = m[0] ?? ''
    if (tok.startsWith('`')) out.push(<code key={key++}>{tok.slice(1, -1)}</code>)
    else if (tok.startsWith('**')) out.push(<strong key={key++}>{tok.slice(2, -2)}</strong>)
    else if (tok.startsWith('[')) {
      const link = /^\[([^\]]+)\]\(([^)\s]+)\)$/.exec(tok)
      if (link)
        out.push(
          <a key={key++} href={link[2]} target="_blank" rel="noreferrer" style={{ color: 'var(--color-brand-secondary, #0f7fff)', textDecoration: 'underline' }}>
            {link[1]}
          </a>,
        )
      else out.push(tok)
    } else out.push(<em key={key++}>{tok.slice(1, -1)}</em>)
    last = i + tok.length
  }
  if (last < text.length) out.push(text.slice(last))
  return out
}

type MdBlock =
  | { kind: 'p'; lines: string[] }
  | { kind: 'ul'; items: string[] }
  | { kind: 'ol'; items: string[] }
  | { kind: 'h'; text: string }
  | { kind: 'code'; text: string }
  | { kind: 'table'; header: string[]; rows: string[][] }
  | { kind: 'hr' }

/** A table separator row: | --- | :--: | ---: | */
function isTableSep(line: string): boolean {
  const t = line.trim()
  return /\|/.test(t) && /^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$/.test(t)
}

function splitRow(line: string): string[] {
  let s = line.trim()
  if (s.startsWith('|')) s = s.slice(1)
  if (s.endsWith('|')) s = s.slice(0, -1)
  return s.split('|').map((c) => c.trim())
}

function parseBlocks(text: string): MdBlock[] {
  const lines = text.split('\n')
  const blocks: MdBlock[] = []
  let cur: MdBlock | null = null
  const flush = (): void => {
    if (cur) {
      blocks.push(cur)
      cur = null
    }
  }
  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i] ?? ''
    const line = raw.trimEnd()
    const t = line.trim()
    // fenced code block
    if (t.startsWith('```')) {
      flush()
      const buf: string[] = []
      i++
      while (i < lines.length && !(lines[i] ?? '').trim().startsWith('```')) {
        buf.push(lines[i] ?? '')
        i++
      }
      blocks.push({ kind: 'code', text: buf.join('\n') })
      continue
    }
    // pipe table: header row + separator row
    if (t.startsWith('|') && i + 1 < lines.length && isTableSep(lines[i + 1] ?? '')) {
      flush()
      const header = splitRow(t)
      i += 2
      const rows: string[][] = []
      while (i < lines.length && (lines[i] ?? '').trim().startsWith('|')) {
        rows.push(splitRow(lines[i] ?? ''))
        i++
      }
      i--
      blocks.push({ kind: 'table', header, rows })
      continue
    }
    if (!t) {
      flush()
      continue
    }
    if (/^(-{3,}|\*{3,}|_{3,})$/.test(t)) {
      flush()
      blocks.push({ kind: 'hr' })
      continue
    }
    const h = /^#{1,6}\s+(.*)$/.exec(line)
    if (h) {
      flush()
      blocks.push({ kind: 'h', text: h[1] ?? '' })
      continue
    }
    const ul = /^\s*[-*•]\s+(.*)$/.exec(line)
    if (ul) {
      if (cur?.kind !== 'ul') {
        flush()
        cur = { kind: 'ul', items: [] }
      }
      cur.items.push(ul[1] ?? '')
      continue
    }
    const ol = /^\s*\d+[.、)]\s+(.*)$/.exec(line)
    if (ol) {
      if (cur?.kind !== 'ol') {
        flush()
        cur = { kind: 'ol', items: [] }
      }
      cur.items.push(ol[1] ?? '')
      continue
    }
    if (cur?.kind !== 'p') {
      flush()
      cur = { kind: 'p', lines: [] }
    }
    cur.lines.push(line)
  }
  flush()
  return blocks
}

const cellStyle: CSSProperties = {
  border: '1px solid var(--color-border-default, #d0d7de)',
  padding: '4px 8px',
  textAlign: 'left',
  verticalAlign: 'top',
}
const thStyle: CSSProperties = {
  ...cellStyle,
  background: 'var(--color-bg-subtle, #f6f8fa)',
  fontWeight: 700,
}

export function Markdown({ text }: { text: string }): React.JSX.Element {
  return (
    <div className="ai-md">
      {parseBlocks(text).map((b, i) => {
        if (b.kind === 'h') {
          return (
            <p key={i} className="ai-md-h">
              {renderInline(b.text)}
            </p>
          )
        }
        if (b.kind === 'ul' || b.kind === 'ol') {
          const items = b.items.map((it, j) => <li key={j}>{renderInline(it)}</li>)
          return b.kind === 'ul' ? <ul key={i}>{items}</ul> : <ol key={i}>{items}</ol>
        }
        if (b.kind === 'code') {
          return (
            <pre
              key={i}
              style={{
                background: 'var(--color-bg-subtle, #f6f8fa)',
                padding: '8px 10px',
                borderRadius: 6,
                overflowX: 'auto',
                margin: '0 0 12px',
                fontSize: 13,
                lineHeight: 1.5,
                fontFamily: "ui-monospace, SFMono-Regular, 'SF Mono', Menlo, Consolas, monospace",
                whiteSpace: 'pre',
              }}
            >
              <code style={{ background: 'none', padding: 0, color: 'inherit', fontSize: 'inherit' }}>
                {b.text}
              </code>
            </pre>
          )
        }
        if (b.kind === 'hr') {
          return (
            <hr
              key={i}
              style={{ border: 'none', borderTop: '1px solid var(--color-border-default, #d0d7de)', margin: '12px 0' }}
            />
          )
        }
        if (b.kind === 'table') {
          return (
            <div key={i} style={{ overflowX: 'auto', margin: '0 0 12px' }}>
              <table style={{ borderCollapse: 'collapse', fontSize: '0.95em', width: '100%' }}>
                <thead>
                  <tr>
                    {b.header.map((c, j) => (
                      <th key={j} style={thStyle}>
                        {renderInline(c)}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {b.rows.map((row, r) => (
                    <tr key={r}>
                      {row.map((c, j) => (
                        <td key={j} style={cellStyle}>
                          {renderInline(c)}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )
        }
        return (
          <p key={i}>
            {b.lines.map((ln, j) => (
              <Fragment key={j}>
                {j > 0 && <br />}
                {renderInline(ln)}
              </Fragment>
            ))}
          </p>
        )
      })}
    </div>
  )
}
