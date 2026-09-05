import { describe, expect, it } from 'vitest'
import type { ChangePlan } from '../src/domain/workbook.types'
import { beautifyTablePlan } from '../src/renderer/plan-operations'

function plan(cellChanges: ChangePlan['cellChanges'], formatChanges: ChangePlan['formatChanges'] = []): ChangePlan {
  return {
    transactionId: 't',
    baseRevision: 0,
    cellChanges,
    sheetRenames: [],
    structuralChanges: [],
    formatChanges,
    warnings: [],
  }
}

const cell = (address: string, value: string | number = 'x', formula?: string) => ({
  sheetId: 's1',
  address,
  before: { value: null },
  after: formula ? { value: null, formula } : { value },
})

describe('beautifyTablePlan', () => {
  it('adds header + border + zebra styling to a fresh table block', () => {
    const out = beautifyTablePlan(
      plan([
        cell('A1', 'No'),
        cell('B1', 'Klausul'),
        cell('A2', '1'),
        cell('B2', '5.1'),
        cell('A3', '2'),
        cell('B3', '5.2'),
      ]),
    )
    const labels = out.formatChanges.map((f) => f.label)
    expect(labels).toContain('Auto-format table header')
    expect(labels).toContain('Auto-format table borders')
    expect(labels).toContain('Auto-format zebra row')
    const header = out.formatChanges.find((f) => f.label === 'Auto-format table header')!
    expect(header.range).toBe('A1:B1')
    expect(header.format.bold).toBe(true)
    expect(header.format.fillColor).toBe('44546A')
    expect(header.format.fontColor).toBe('FFFFFF')
  })

  it('does not touch a sheet that already has explicit formatting', () => {
    const out = beautifyTablePlan(
      plan(
        [cell('A1'), cell('B1'), cell('A2'), cell('B2')],
        [{ sheetId: 's1', range: 'A1:B1', format: { bold: true }, label: 'manual' }],
      ),
    )
    expect(out.formatChanges.map((f) => f.label)).toEqual(['manual'])
  })

  it('ignores single-row or single-column writes (not a table)', () => {
    const singleRow = beautifyTablePlan(plan([cell('A1'), cell('B1')]))
    expect(singleRow.formatChanges).toHaveLength(0)
    const singleCol = beautifyTablePlan(plan([cell('A1'), cell('A2')]))
    expect(singleCol.formatChanges).toHaveLength(0)
  })

  it('never shifts rows/columns (formula-safe)', () => {
    const out = beautifyTablePlan(
      plan([
        cell('A1', 'No'),
        cell('B1', 'Total'),
        cell('A2', '1'),
        cell('B2', '=SUM(C2:C10)', '=SUM(C2:C10)'),
        cell('A3', '2'),
        cell('B3', '=B2*2', '=B2*2'),
      ]),
    )
    // Only formatChanges are added; cellChanges (incl. formulas) are untouched.
    expect(out.cellChanges).toHaveLength(6)
    expect(out.cellChanges.some((c) => c.after.formula)).toBe(true)
    expect(out.formatChanges.length).toBeGreaterThan(0)
  })
})
