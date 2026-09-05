import type { AgentSkill, AgentToolCall, AgentToolDef, ToolExecution } from '@prova/agent-core'
import {
  NODE_TYPE_LABEL,
  nextId,
  autoLayout,
  autoLayoutWithLanes,
  computeLaneLayout,
  validateFlow,
} from '../flow-model'
import {
  type FlowDocument,
  type FlowEdge,
  type FlowLane,
  type FlowNode,
  type NodeType,
} from '../../shared/ipc'

const VALID_TYPES: NodeType[] = ['start', 'end', 'process', 'decision', 'io']

function isNodeType(v: unknown): v is NodeType {
  return typeof v === 'string' && (VALID_TYPES as string[]).includes(v)
}

const SYSTEM_PROMPT = `You are the AI assistant inside PROVAOffice Flows, a flowchart editor. You build and edit a flowchart through tools.

The diagram is a set of nodes and directed edges:
- node types: start, end, process, decision, io
- decision nodes usually have 2+ outgoing edges with yes/no or condition labels

## Creating a whole flow
PLAN FIRST: before calling any tool, write a short plan in your reply text: the main happy-path steps top-to-bottom, where (if anywhere) a decision is truly needed, and who owns each step (lanes). Then call generate_flow ONCE with the COMPLETE plan - do not build piecemeal with many small add_* calls.
generate_flow plan: { title, lanes:[{id,name}], nodes:[{id,type,label,lane}], edges:[{id,from,to,label?}] }.
- SWIMLANES: lanes = process owners/roles (e.g. Customer, Manager, HR, System); put each node in the lane of whoever performs it. Use short lane ids like l1,l2. Omit lanes if no distinct owners.
- Use short stable ids like n1,n2 and e1,e2.
- Give every process/decision/io node a concise label (match the user's language).
- Exactly one 'start' node; usually one 'end'.
- FLOWCHART LOGIC (strict): every 'decision' MUST have >=2 outgoing edges with labels (yes/no or conditions). Every 'process'/'io' MUST have AT MOST 1 outgoing edge. 'start' has exactly 1; 'end' has 0. The tool rejects plans that break these.
- Do NOT set coordinates - layout is computed automatically.
- Do NOT loop web_search; if you need domain facts, do at most 1-2 searches first, then generate_flow.

## Editing an existing flow
Use the fine-grained tools: add_node, add_edge, delete_node, set_label, move_node, auto_layout.
Reference nodes by the ids shown in the current-diagram context.

## Keep it simple and readable (IMPORTANT)
- Model the MAIN / happy-path flow first as a clean top-to-bottom sequence.
- Use a 'decision' (diamond) ONLY when a real branch is needed. Prefer linear process steps otherwise.
- Keep the number of decisions low (usually 0-2). Avoid nested/extra decisions that clutter the diagram.
- Flow must go top-to-bottom: avoid back-edges (an edge from a lower node to a higher node). If a loop-back is truly required, keep it to a single clear case.
- Keep the graph mostly linear so the logic is easy to follow.

## Refine loop (IMPORTANT)
After building or editing the flow, ALWAYS run this loop until it is clean:
1. generate_flow (or an edit tool)
2. ALWAYS call validate_flow immediately after generate_flow (never skip it)
3. If validate_flow reports problems, fix them with the edit tools (add_edge, set_label, delete_node, add_node) - e.g. add the missing labeled branch on a decision, or connect an unreachable node.
4. validate_flow again. Repeat until it returns OK, up to 5-7 iterations.
Only give your final summary once validate_flow returns OK. Never stop while known problems remain.

## Rules
- After any change, the diagram re-lays-out automatically.
- Keep labels short so they fit the shapes.
- Finish with a one-sentence plain-text summary (no markdown tables).`

function toolDefs(): AgentToolDef[] {
  return [
    {
      name: 'read_flow',
      description: 'Return the current diagram (nodes + edges) as JSON.',
      inputSchema: { type: 'object', properties: {}, required: [] },
    },
    {
      name: 'generate_flow',
      description:
        'Replace the whole diagram with a complete plan. Provide lanes (id,name) for process owners, nodes (id,type,label,lane) and edges (id,from,to,label?). Coordinates are computed automatically. Use lanes to show WHO owns each step (swimlanes).',
      inputSchema: {
        type: 'object',
        properties: {
          title: { type: 'string' },
          lanes: {
            type: 'array',
            description: 'Swimlane columns = process owners/roles. Order = left to right.',
            items: {
              type: 'object',
              properties: { id: { type: 'string' }, name: { type: 'string' } },
              required: ['id', 'name'],
            },
          },
          nodes: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                id: { type: 'string' },
                type: { type: 'string', enum: VALID_TYPES },
                label: { type: 'string' },
                lane: { type: 'string', description: 'id of the lane (process owner) this node belongs to' },
              },
              required: ['id', 'type', 'label'],
            },
          },
          edges: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                id: { type: 'string' },
                from: { type: 'string' },
                to: { type: 'string' },
                label: { type: 'string' },
              },
              required: ['id', 'from', 'to'],
            },
          },
        },
        required: ['nodes', 'edges'],
      },
    },
    {
      name: 'add_node',
      description: 'Add one node to the current diagram.',
      inputSchema: {
        type: 'object',
        properties: {
          type: { type: 'string', enum: VALID_TYPES },
          label: { type: 'string' },
          id: { type: 'string', description: 'optional id; auto-generated if omitted' },
        },
        required: ['type', 'label'],
      },
    },
    {
      name: 'add_edge',
      description: 'Connect two existing nodes with a directed edge.',
      inputSchema: {
        type: 'object',
        properties: {
          from: { type: 'string' },
          to: { type: 'string' },
          label: { type: 'string' },
        },
        required: ['from', 'to'],
      },
    },
    {
      name: 'delete_node',
      description: 'Remove a node and its connected edges.',
      inputSchema: {
        type: 'object',
        properties: { id: { type: 'string' } },
        required: ['id'],
      },
    },
    {
      name: 'set_label',
      description: 'Change the label text of a node or edge.',
      inputSchema: {
        type: 'object',
        properties: {
          id: { type: 'string' },
          label: { type: 'string' },
        },
        required: ['id', 'label'],
      },
    },
    {
      name: 'auto_layout',
      description: 'Re-run automatic layered layout to tidy the diagram.',
      inputSchema: { type: 'object', properties: {}, required: [] },
    },
    {
      name: 'validate_flow',
      description: 'Check the current diagram against flowchart rules (single start, >=1 end, every decision has >=2 LABELED outgoing edges, process/io at most 1 output, no unreachable nodes). Call this after generate_flow and after edits; fix any reported problems and validate again until it returns OK.',
      inputSchema: { type: 'object', properties: {}, required: [] },
    },
  ]
}

export interface FlowsSkillAccess {
  getDoc(): FlowDocument
  setDoc(doc: FlowDocument): void
}

export function createFlowsSkill(access: FlowsSkillAccess): AgentSkill {
  const findNode = (id: string): FlowNode | undefined => access.getDoc().nodes.find((n) => n.id === id)

  return {
    id: 'flows',
    systemPrompt: SYSTEM_PROMPT,
    tools: toolDefs(),
    buildContext: () => {
      const d = access.getDoc()
      return `<current_diagram>\n${JSON.stringify({ title: d.title, lanes: d.lanes ?? [], nodes: d.nodes.map((n) => ({ id: n.id, type: n.type, label: n.label, lane: n.lane })), edges: d.edges.map((e) => ({ id: e.id, from: e.from, to: e.to, label: e.label })) })}\n</current_diagram>`
    },
    executeTool: async (call: AgentToolCall): Promise<ToolExecution> => {
      const d = access.getDoc()
      switch (call.name) {
        case 'read_flow':
          return { output: JSON.stringify(d), summary: 'Read diagram' }

        case 'generate_flow': {
          const nodesIn = Array.isArray(call.input.nodes) ? (call.input.nodes as unknown[]) : []
          const edgesIn = Array.isArray(call.input.edges) ? (call.input.edges as unknown[]) : []
          const lanesIn = Array.isArray(call.input.lanes) ? (call.input.lanes as unknown[]) : []
          const lanes: FlowLane[] = []
          for (const raw of lanesIn) {
            const o = (raw ?? {}) as Record<string, unknown>
            const id = String(o.id ?? nextId('lane'))
            if (lanes.some((l) => l.id === id)) continue
            lanes.push({ id, name: String(o.name ?? id) })
          }
          const nodes: FlowNode[] = []
          for (const raw of nodesIn) {
            const o = (raw ?? {}) as Record<string, unknown>
            const type = isNodeType(o.type) ? o.type : 'process'
            const id = String(o.id ?? nextId('n'))
            if (nodes.some((n) => n.id === id)) continue
            nodes.push({ id, type, label: String(o.label ?? NODE_TYPE_LABEL[type]), x: 0, y: 0, lane: o.lane ? String(o.lane) : undefined })
          }
          const ids = new Set(nodes.map((n) => n.id))
          const edges: FlowEdge[] = []
          for (const raw of edgesIn) {
            const o = (raw ?? {}) as Record<string, unknown>
            const from = String(o.from ?? '')
            const to = String(o.to ?? '')
            if (!ids.has(from) || !ids.has(to)) continue
            edges.push({
              id: String(o.id ?? nextId('e')),
              from,
              to,
              label: o.label ? String(o.label) : undefined,
            })
          }
          // Flowchart logic validation: decisions branch, processes do not.
          const outCount = new Map<string, number>()
          for (const e of edges) outCount.set(e.from, (outCount.get(e.from) ?? 0) + 1)
          const problems: string[] = []
          for (const n of nodes) {
            const out = outCount.get(n.id) ?? 0
            if (n.type === 'decision' && out < 2) problems.push(`decision "${n.id}" (${n.label}) must have >=2 outgoing edges (e.g. yes/no)`)
            if ((n.type === 'process' || n.type === 'io') && out > 1) problems.push(`${n.type} "${n.id}" (${n.label}) must have at most 1 outgoing edge`)
            if (n.type === 'start' && out !== 1) problems.push(`start "${n.id}" must have exactly 1 outgoing edge`)
            if (n.type === 'end' && out > 0) problems.push(`end "${n.id}" must have no outgoing edges`)
          }
          if (problems.length) {
            return {
              output: `Flowchart logic errors (fix and call generate_flow again): ${problems.join('; ')}`,
              isError: true,
              summary: 'Invalid flowchart logic',
            }
          }

          const baseDoc: FlowDocument = {
            version: 1,
            title: String(call.input.title ?? d.title),
            lanes: lanes.length ? lanes : undefined,
            nodes,
            edges,
          }
          const laid = lanes.length ? autoLayoutWithLanes(baseDoc) : await autoLayout(baseDoc)
          access.setDoc(laid)
          return {
            output: `Generated ${laid.nodes.length} nodes, ${laid.edges.length} edges.`,
            mutated: true,
            summary: `Generated flow (${laid.nodes.length} nodes)`,
          }
        }

        case 'add_node': {
          const type = isNodeType(call.input.type) ? call.input.type : 'process'
          const id = String(call.input.id ?? nextId('n'))
          if (findNode(id)) return { output: `id ${id} exists`, isError: true, summary: 'Add node' }
          const node: FlowNode = {
            id,
            type,
            label: String(call.input.label ?? NODE_TYPE_LABEL[type]),
            x: 0,
            y: 0,
          }
          const next = { ...d, nodes: [...d.nodes, node] }
      const laid = next.lanes && next.lanes.length ? autoLayoutWithLanes(next) : await autoLayout(next)
          access.setDoc(laid)
          return { output: `Added node ${id}`, mutated: true, summary: `Added ${type} node` }
        }

        case 'add_edge': {
          const from = String(call.input.from ?? '')
          const to = String(call.input.to ?? '')
          if (!findNode(from) || !findNode(to))
            return { output: 'unknown from/to node id', isError: true, summary: 'Add edge' }
          const edge: FlowEdge = {
            id: nextId('e'),
            from,
            to,
            label: call.input.label ? String(call.input.label) : undefined,
          }
          access.setDoc({ ...d, edges: [...d.edges, edge] })
          return { output: `Added edge ${edge.id}`, mutated: true, summary: 'Connected nodes' }
        }

        case 'delete_node': {
          const id = String(call.input.id ?? '')
          if (!findNode(id)) return { output: `no node ${id}`, isError: true, summary: 'Delete node' }
          access.setDoc({
            ...d,
            nodes: d.nodes.filter((n) => n.id !== id),
            edges: d.edges.filter((e) => e.from !== id && e.to !== id),
          })
          return { output: `Deleted ${id}`, mutated: true, summary: 'Deleted node' }
        }

        case 'set_label': {
          const id = String(call.input.id ?? '')
          const label = String(call.input.label ?? '')
          const node = findNode(id)
          if (node) {
            access.setDoc({
              ...d,
              nodes: d.nodes.map((n) => (n.id === id ? { ...n, label } : n)),
            })
            return { output: `Renamed node ${id}`, mutated: true, summary: 'Updated label' }
          }
          if (d.edges.some((e) => e.id === id)) {
            access.setDoc({
              ...d,
              edges: d.edges.map((e) => (e.id === id ? { ...e, label } : e)),
            })
            return { output: `Renamed edge ${id}`, mutated: true, summary: 'Updated label' }
          }
          return { output: `no element ${id}`, isError: true, summary: 'Set label' }
        }

        case 'validate_flow': {
          const problems = validateFlow(d)
          if (problems.length === 0)
            return { output: 'OK - the flowchart is logically valid.', summary: 'Validation passed' }
          const lines = problems.map((p, i) => `${i + 1}. ${p}`).join('\n')
          return {
            output: `Found ${problems.length} problem(s):` + "\n" + lines + "\n" + 'Fix these with the edit tools (or regenerate), then call validate_flow again.',
            summary: `${problems.length} logic issue(s)`,
          }
        }
        case 'auto_layout': {
          const laid = d.lanes && d.lanes.length ? autoLayoutWithLanes(d) : await autoLayout(d)
          access.setDoc(laid)
          return { output: 'Re-laid out', mutated: true, summary: 'Auto layout' }
        }

        default:
          return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
    },
  }
}
