/**
 * Second round of the QA pass of 18/11/2026 (findings rated medium and low; ids as in its report): recalculation,
 * what is typed, sorting, formats of whole columns, collaboration.
 */
import { describe, expect, test } from 'bun:test'
import { at, book, openBook, rows, type Opened } from './xlsx-builder'

const recalculated = async (o: Opened) => {
  while (!o.events.some((e) => e.event === 'xlsx:progress' && e.data.recalc)) await new Promise((r) => setTimeout(r, 5))
}
const num = async (o: Opened, sheet: number, ref: string) => Number(((await o.engine.call('xlsx.cell', [sheet, at(ref).r, at(ref).c])) as any)?.value)

describe('P1-013: a sheet where every row reads the whole table', () => {
  const N = 6000
  test('running totals, COUNTIF and SUMIFS down a column follow one edit, with the right values', async () => {
    const grid: any[][] = []
    for (let r = 1; r <= N; r++) {
      grid.push([r % 50, r % 7, r, { f: `SUM($C$1:C${r})`, v: 0 }, { f: `COUNTIF(A:A,A${r})`, v: 0 }, { f: `SUMIFS(C:C,A:A,A${r},B:B,B${r})`, v: 0 }, { f: `VLOOKUP(${(r % 40) + 1},$A$1:$C$60,3,FALSE)`, v: 0 }])
    }
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows(grid) }] }))
    await o.set(0, 'Z1', '1')
    const t0 = performance.now()
    await o.set(0, 'C1', '1000')
    await o.set(0, 'A2', '49')
    const ms = performance.now() - t0
    // by hand
    const A = grid.map((g) => g[0] as number)
    const B = grid.map((g) => g[1] as number)
    const C = grid.map((g) => g[2] as number)
    C[0] = 1000
    A[1] = 49
    const sumifs = (r: number) => C.reduce((s, c, i) => (A[i] === A[r] && B[i] === B[r] ? s + c : s), 0)
    for (const r of [0, 1, 2, 48, 49, 349, N - 1]) {
      expect(await num(o, 0, `D${r + 1}`)).toBe(C.slice(0, r + 1).reduce((x, y) => x + y, 0))
      expect(await num(o, 0, `E${r + 1}`)).toBe(A.filter((a) => a === A[r]).length)
      expect(await num(o, 0, `F${r + 1}`)).toBe(sumifs(r))
      expect(await num(o, 0, `G${r + 1}`)).toBe(C[A.slice(0, 60).indexOf(((r + 1) % 40) + 1)])
    }
    // the cells were read twenty million times per column: seconds, not minutes (the bound is for a slow machine)
    expect(ms).toBeLessThan(60_000)
  }, 120_000)

  test('a block kept while computing is read again once a formula in it has its value', async () => {
    // B reads A whole, twice over (kept from the second time); A2 is a formula still to compute when B1 reads it
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1, { f: 'SUM($A$1:$A$3)', v: 0 }], [{ f: 'A1*10', v: 0 }, { f: 'SUM($A$1:$A$3)', v: 0 }], [3, { f: 'SUM($A$1:$A$3)+C3', v: 0 }, { f: 'A2+1', v: 0 }]]) }] }))
    await o.set(0, 'A1', '2')
    expect([await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 1)]).toEqual(['25', '25', '46'])
    await o.set(0, 'A3', '=B1')
    // a circular reference through a kept block: an error, the same however often it is computed
    const first = [await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 0)]
    await o.set(0, 'A1', '2')
    expect([await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 0)]).toEqual(first)
    expect(first).toContain('#REF!')
    await o.set(0, 'A3', '4')
    expect([await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 1)]).toEqual(['26', '26', '47'])
  })
})

describe('P1-025: TODAY, NOW and RAND are of the moment the file is read', () => {
  test('what the file holds for them, and for what reads them, is computed again; the file is not changed by it', async () => {
    const src = await book({
      sheets: [
        { name: 'Due', rows: rows([[{ f: 'TODAY()', v: 45000 }, 44990, { f: 'A1-B1', v: 10 }, { f: 'NOW()', v: 45000.5 }, { f: 'RAND()', v: 0.123 }, { f: 'B1*2', v: 1 }]]) },
        { name: 'Other', rows: rows([[{ f: 'Due!C1+1', v: 11 }]]) },
      ],
    })
    const o = await openBook(src)
    await recalculated(o)
    const today = await num(o, 0, 'A1')
    expect(today).toBeGreaterThan(46000)
    expect(Number.isInteger(today)).toBe(true)
    expect(await num(o, 0, 'C1')).toBe(today - 44990)
    expect(await num(o, 0, 'D1')).toBeGreaterThanOrEqual(today)
    expect(await num(o, 0, 'E1')).not.toBe(0.123)
    // what does not read the clock keeps the value of the file, right or not: nothing asked to compute it
    expect(await num(o, 0, 'F1')).toBe(1)
    await o.load(1)
    expect(await num(o, 1, 'A1')).toBe(today - 44990 + 1)
    expect(await o.engine.call('dirty', [])).toBe(false)
  })
})

describe('what is typed (P1-027, N1-007, N1-022)', () => {
  const STYLES = (fmts: string[]) =>
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="${fmts.length}">${fmts.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${f}"/>`).join('')}</numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${fmts.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>${fmts.map((_, i) => `<xf numFmtId="${164 + i}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join('')}</cellXfs></styleSheet>`
  const cell = async (o: Opened, ref: string) => (await o.engine.call('xlsx.cell', [0, at(ref).r, at(ref).c])) as any

  test('"+A1+B1", "-A1" and "@SUM(…)" are formulas; "-5", "- item" and a phone number are not', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'D', rows: rows([[1, 2, 3]]) }] }))
    await o.set(0, 'A3', '+A1+B1')
    await o.set(0, 'B3', '-C1')
    await o.set(0, 'C3', '@SUM(A1:C1)')
    await o.set(0, 'D3', '+D!A1+D!B1')
    expect([await o.formula(0, 2, 0), await o.shown(0, 2, 0)]).toEqual(['=+A1+B1', '3'])
    expect([await o.formula(0, 2, 1), await o.shown(0, 2, 1)]).toEqual(['=-C1', '-3'])
    expect([await o.formula(0, 2, 2), await o.shown(0, 2, 2)]).toEqual(['=SUM(A1:C1)', '6'])
    expect(await o.shown(0, 2, 3)).toBe('3')
    await o.set(0, 'A5', '-5')
    await o.set(0, 'B5', '- item')
    await o.set(0, 'C5', '+84 912 345 678')
    await o.set(0, 'D5', '@username')
    await o.set(0, 'E5', '-abc')
    expect((await cell(o, 'A5')).type).toBe('number')
    for (const ref of ['B5', 'C5', 'D5', 'E5']) expect([(await cell(o, ref)).type, (await cell(o, ref)).formula]).toEqual(['string', undefined])
  })

  test('an error typed is an error; a number longer than a number holds stays as typed; amounts are numbers shown as typed', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    await o.set(0, 'A2', '#N/A')
    await o.set(0, 'B2', '=ISNA(A2)')
    await o.set(0, 'C2', '=IFERROR(A2,"none")')
    expect([(await cell(o, 'A2')).type, await o.shown(0, 1, 1), await o.shown(0, 1, 2)]).toEqual(['error', 'TRUE', 'none'])
    await o.set(0, 'A3', '9704123456789012345')
    expect([(await cell(o, 'A3')).type, await o.shown(0, 2, 0)]).toEqual(['string', '9704123456789012345'])
    await o.set(0, 'B3', '123456789012345')
    expect((await cell(o, 'B3')).type).toBe('number')
    const typed: [string, string, number, string][] = [
      ['A4', '1,234,567', 1234567, '1,234,567'],
      ['B4', '$50', 50, '$50'],
      ['C4', '€1,200.5', 1200.5, '€1,200.50'],
      ['D4', '50 ₫', 50, '50 ₫'],
      ['E4', '-$3.75', -3.75, '-$3.75'],
      ['F4', '1,234.5', 1234.5, '1,234.5'],
    ]
    for (const [ref, input, n, shown] of typed) {
      await o.set(0, ref, input)
      expect([ref, Number((await cell(o, ref)).value), await o.shown(0, at(ref).r, at(ref).c)]).toEqual([ref, n, shown])
    }
    // not amounts: an address, a decimal comma, two signs
    for (const [ref, input] of [['A6', '192.168.100.200'], ['B6', '1,5'], ['C6', '$5€'], ['D6', '12,34,567']] as const) {
      await o.set(0, ref, input)
      expect([ref, (await cell(o, ref)).type]).toEqual([ref, 'string'])
    }
    const re = await openBook((await o.save()).blob)
    expect([await re.shown(0, 3, 1), await re.shown(0, 3, 2), await re.shown(0, 1, 0)]).toEqual(['$50', '€1,200.50', '#N/A'])
  })

  test('a formula typed into a cell without a format shows as what it computes', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[{ s: 1, v: 45292 }, { s: 2, v: 0.25 }, 5, { s: 1, v: 45300 }]]) }], styles: STYLES(['dd/mm/yyyy', '0.0%']) }))
    expect(await o.shown(0, 0, 0)).toBe('01/01/2024')
    await o.set(0, 'A2', '=A1')
    await o.set(0, 'B2', '=A1+30')
    await o.set(0, 'C2', '=DATE(2024,2,29)')
    await o.set(0, 'D2', '=EDATE(A1,1)')
    await o.set(0, 'E2', '=B1')
    await o.set(0, 'F2', '=D1-A1')
    await o.set(0, 'G2', '=C1+1')
    await o.set(0, 'H2', '=MAX(A1:D1)')
    expect([await o.shown(0, 1, 0), await o.shown(0, 1, 1), await o.shown(0, 1, 2), await o.shown(0, 1, 3)]).toEqual(['01/01/2024', '31/01/2024', '29/02/2024', '01/02/2024'])
    expect([await o.shown(0, 1, 4), await o.shown(0, 1, 5), await o.shown(0, 1, 6), await o.shown(0, 1, 7)]).toEqual(['25.0%', '8', '6', '09/01/2024'])
    await o.set(0, 'A3', '=TODAY()')
    expect(await o.shown(0, 2, 0)).toMatch(/^\d{2}\/\d{2}\/20\d{2}$/)
    // a cell that has a format keeps it
    await o.set(0, 'B1', '=A1')
    expect(await o.shown(0, 0, 1)).toBe('4529200.0%')
    // undone with its format
    await o.set(0, 'J2', '=A1')
    await o.engine.call('xlsx.undo', [])
    await o.set(0, 'J2', '7')
    expect(await o.shown(0, 1, 9)).toBe('7')
  })
})
